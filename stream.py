import streamlit as st
import tempfile
import shutil
import time
import os

from sentence_transformers import SentenceTransformer
from text_processing import TextProcessor
from config import file_path_dictionnary, file_path_json, MODEL_DIR_1
from sentenceT import process_new_descriptions

st.set_page_config(page_title="Classificateur Takalot", layout="centered")
st.title("📋 Classificateur de défauts de Cabine")

uploaded_file = st.file_uploader(
    "Glissez ici un fichier Excel à traiter", type=["xlsx", "xlsm"]
)

if uploaded_file:
    st.success(f"✅ Fichier reçu : {uploaded_file.name}")

    if st.button("🚀 Lancer le traitement"):
        progress = st.progress(0, text="Initialisation du traitement...")
        time.sleep(0.5)

        # On crée un fichier temporaire avec suffixe .xlsx
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_file.read())
            temp_input_path = tmp.name

        model = SentenceTransformer(MODEL_DIR_1)
        processor = TextProcessor(file_path_dictionnary)

        try:
            start_time = time.time()
            progress.progress(20, text="🧠 Chargement du modèle et des données...")

            process_new_descriptions(
                file_path_takalot_file=temp_input_path,
                embedding_model=model,
                processor=processor,
                reference_json=file_path_json
            )

            progress.progress(90, text="💾 Sauvegarde du fichier modifié...")

            final_path = "Data_Cabine/Output/fichier_modifie.xlsx"
            os.makedirs(os.path.dirname(final_path), exist_ok=True)
            shutil.copy(temp_input_path, final_path)

            progress.progress(100, text="✅ Terminé")
            elapsed = time.time() - start_time
            st.success(f"🎉 Traitement terminé en {elapsed:.2f} secondes")
            st.info(f"📁 Le fichier modifié a été sauvegardé ici : `{final_path}`")

        except Exception as e:
            st.error(f"❌ Erreur pendant le traitement : {e}")
