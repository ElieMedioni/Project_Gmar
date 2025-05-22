from openpyxl import load_workbook, Workbook
import numpy as np
import json
from scipy.spatial.distance import cdist
import datetime

import streamlit as st


def load_embeddings_from_json(json_path):
    with open(json_path, "r", encoding="utf-8") as f:
        return json.load(f)

def process_new_descriptions(file_path_takalot_file, embedding_model, processor, reference_json, log_fn=print):
    
    log_fn("🧽 Nettoyage du fichier Excel d'origine...")

    # 1. Charger le fichier d'origine en mode valeur uniquement
    wb_src = load_workbook(file_path_takalot_file, data_only=True)
    ws_src = wb_src.active

    # 4. Créer un nouveau classeur propre
    wb_clean = Workbook()
    ws_clean = wb_clean.active
                
    for row in ws_src.iter_rows():
        has_value = any(cell.value is not None and str(cell.value).strip() != "" for cell in row)
        if has_value:
            for cell in row:
                value = cell.value
                if isinstance(value, datetime.datetime):
                    value = value.date()
                ws_clean.cell(row=cell.row, column=cell.column, value=value)

    # 6. Sauvegarder le fichier nettoyé
    wb_clean.save(file_path_takalot_file)
    log_fn(f"✅ Fichier nettoyé sauvegardé dans : {file_path_takalot_file}")

    # 7. Recharger ce fichier nettoyé pour traitement
    log_fn("📂 Ouverture du fichier Excel nettoyé...")
    wb = load_workbook(file_path_takalot_file)
    ws = wb.active
    log_fn(f"✅ Accès à la feuille : {ws.title}")

    # 🔍 Recherche de la cellule contenant "WO Desc"
    desc_cell = None
    for row in ws.iter_rows(min_row=1, max_row=30, max_col=30):
        for cell in row:
            if cell.value:
                value = str(cell.value).strip().lower()
                if value == "wo desc":
                    desc_cell = cell
        if desc_cell:
            break

    if not desc_cell:
        log_fn("❌ Header 'WO Desc' not found")
        return

    log_fn(f"🔎 'WO Desc' found in {desc_cell.coordinate}")

    if desc_cell.row > 1:
        ws.delete_rows(1, desc_cell.row - 1)
        log_fn(f"🧹 Lignes supprimées au-dessus de {desc_cell.coordinate}")
        desc_cell = ws.cell(row=1, column=desc_cell.column)

    if ws.cell(row=1, column=1).value is None or str(ws.cell(row=1, column=1).value).strip() == "":
        ws.delete_cols(1)
        log_fn("🧹 Première colonne supprimée car vide")


    # 🔍 Recherche des colonnes Main/Sub Category
    main_cell = None
    sub_cell = None

    for cell in ws[desc_cell.row]:
        if cell.value:
            value = str(cell.value).strip().lower()
            if value == "main category":
                main_cell = cell
            elif value == "sub category":
                sub_cell = cell

    col_index = desc_cell.column
    start_row = desc_cell.row + 1

    descriptions = []
    row_indexes = []

    for row in range(start_row, ws.max_row + 1):
        value = ws.cell(row=row, column=col_index).value
        if value is not None and str(value).strip():
            descriptions.append(str(value))
            row_indexes.append(row)
            log_fn(f"📖 Ligne {row} lue")
        else:
            break  # on s'arrête dès la première ligne vide

    log_fn("🧼 Nettoyage des descriptions...")
    cleaned = [processor.clean_text(text) for text in descriptions]

    log_fn("🔍 Génération des embeddings...")
    takalot_embeddings = embedding_model.encode(
        cleaned,
        normalize_embeddings=True,
        show_progress_bar=True,
        batch_size=32
    )
    
    log_fn("📥 Chargement des embeddings de référence...")
    reference_embeddings = load_embeddings_from_json(reference_json)
    ref_keys = list(reference_embeddings.keys())
    ref_matrix = np.array([reference_embeddings[k] for k in ref_keys])

    log_fn("📏 Calcul des distances (euclidean)...")
    dist_matrix = cdist(takalot_embeddings, ref_matrix, metric="euclidean")
    top_1_idx = np.argmin(dist_matrix, axis=1)
        
        
    log_fn("✏️ Écriture des résultats dans le fichier Excel...")
    progress_bar = st.progress(0)

    for i, row_num in enumerate(row_indexes):
        best = ref_keys[top_1_idx[i]]
        cat_1, subcat_1 = best.split(" | ")

        ws.cell(row=row_num, column=main_cell.column).value = cat_1
        ws.cell(row=row_num, column=sub_cell.column).value = subcat_1

        progress_percent = int((i + 1) / len(row_indexes) * 100)
        progress_bar.progress(progress_percent)

    # 🧹 Supprimer la barre à la fin
    progress_bar.empty()

    log_fn("💾 Sauvegarde du fichier...")
    wb.save(file_path_takalot_file)
