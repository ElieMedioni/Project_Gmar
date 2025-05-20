"""from openpyxl import load_workbook
import numpy as np
import json
from scipy.spatial.distance import cdist

def load_embeddings_from_json(json_path):
    with open(json_path, "r", encoding="utf-8") as f:
        return json.load(f)

def is_macro_file(filepath):
    return filepath.lower().endswith(".xlsm")

def process_new_descriptions(file_path_takalot_file, embedding_model, processor, reference_json):
    print("📂 Ouverture du fichier Excel avec openpyxl...")

    if is_macro_file(file_path_takalot_file):
        wb = load_workbook(file_path_takalot_file, keep_vba=True)
    else:
        wb = load_workbook(file_path_takalot_file)

    ws = wb.active

    descriptions = []
    row_indexes = []

    for row in ws.iter_rows(min_row=14, min_col=13, max_col=13):  # Colonne M
        cell = row[0]
        if cell.value is not None:
            descriptions.append(str(cell.value))
            row_indexes.append(cell.row)

    print("🧼 Nettoyage des descriptions...")
    cleaned = [processor.clean_text(text) for text in descriptions]

    print("🔍 Génération des embeddings...")
    takalot_embeddings = embedding_model.encode(
        cleaned,
        normalize_embeddings=True,
        show_progress_bar=True,
        batch_size=32
    )

    reference_embeddings = load_embeddings_from_json(reference_json)
    ref_keys = list(reference_embeddings.keys())
    ref_matrix = np.array([reference_embeddings[k] for k in ref_keys])

    print("📏 Calcul des distances (euclidean)...")
    dist_matrix = cdist(takalot_embeddings, ref_matrix, metric="euclidean")
    top_2_idx = np.argsort(dist_matrix, axis=1)[:, :2]

    print("✏️ Ajout des correspondances dans Excel...")
    for i, row_num in enumerate(row_indexes):
        best_1 = ref_keys[top_2_idx[i, 0]]
        cat_1, subcat_1 = best_1.split(" | ")
        ws.cell(row=row_num, column=17).value = cat_1     # Colonne Q
        ws.cell(row=row_num, column=18).value = subcat_1  # Colonne R

    print("💾 Sauvegarde du fichier...")
    wb.save(file_path_takalot_file)
    print(f"✅ Fichier mis à jour avec succès : {file_path_takalot_file}")"""




import xlwings as xw
import numpy as np
import json
from scipy.spatial.distance import cdist

def load_embeddings_from_json(json_path):
    with open(json_path, "r", encoding="utf-8") as f:
        return json.load(f)

def process_new_descriptions(file_path_takalot_file, embedding_model, processor, reference_json, log_fn=print):
    log_fn("📂 Ouverture du fichier Excel avec xlwings...")
    
    try:
        app = xw.App(visible=True)
        log_fn("✅ xlwings lancé")
        wb = app.books.open(file_path_takalot_file)
        log_fn("✅ Fichier Excel ouvert")
    except Exception as e:
        log_fn(f"❌ Erreur d'ouverture Excel : {e}")
        return
    
    ws = wb.sheets[0]
    log_fn(f"✅ Accès à la feuille : {ws.name}")

    descriptions = []
    row_indexes = []

    # Lecture de la colonne M (colonne 13), lignes à partir de 14
    row = 14
    while True:
        value = ws.range(f"M{row}").value
        if value is None or str(value).strip() == "":  # si la cellule est vide, on arrête la boucle
            break
        descriptions.append(str(value))
        row_indexes.append(row)
        log_fn(f"📖 Ligne {row} lues")
        row += 1


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
    top_2_idx = np.argsort(dist_matrix, axis=1)[:, :2]

    log_fn("✏️ Écriture des résultats dans le fichier Excel...")
    for i, row in enumerate(row_indexes):
        best_1 = ref_keys[top_2_idx[i, 0]]
        cat_1, subcat_1 = best_1.split(" | ")
        ws.range(f"Q{row}").value = cat_1
        ws.range(f"R{row}").value = subcat_1
        log_fn(f"📖 Ligne {row} ecrites")

    print("💾 Sauvegarde du fichier (slicers conservés)...")
    wb.save()
    wb.close()
    app.quit()
    print(f"✅ Fichier mis à jour avec succès : {file_path_takalot_file}")
