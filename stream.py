import streamlit as st
import tempfile
import shutil
import time
import os

from sentence_transformers import SentenceTransformer
from text_processing import TextProcessor
from config import file_path_dictionnary, file_path_json, MODEL_DIR_1
from sentenceT import process_new_descriptions

from data_loader import CategoryEmbeddingBuilder
from config import file_path_categories

st.set_page_config(page_title="Classificateur Takalot", layout="centered")
st.title("📋 Classificateur de défauts de Cabine")

##########################################################################

if st.button("🧠 Générer un nouveau fichier d'embeddings"):
    with st.spinner("Génération des embeddings en cours..."):
        try:
            builder = CategoryEmbeddingBuilder(file_path_categories, MODEL_DIR_1)
            builder.save_outputs()
            st.success("✅ Embeddings créés avec succès et sauvegardés dans `embeddings.json`")
        except Exception as e:
            st.error(f"❌ Erreur lors de la génération des embeddings : {e}")

##########################################################################

def generate_unique_filename(base_path, base_name="fichier_modifie", ext=".xlsx"):
    i = 0
    while True:
        suffix = f"_{i}" if i > 0 else ""
        filename = f"{base_name}{suffix}{ext}"
        full_path = os.path.join(base_path, filename)
        if not os.path.exists(full_path):
            return full_path
        i += 1
        
##########################################################################

uploaded_file = st.file_uploader(
    "Glissez ici un fichier Excel à traiter", type=["xlsx", "xlsm"]
)

if uploaded_file:
    st.success(f"✅ Fichier reçu : {uploaded_file.name}")

    if st.button("🚀 Lancer le traitement"):
        progress = st.progress(0, text="Initialisation du traitement...")
        time.sleep(0.5)

        # 🧠 Déduire l'extension du fichier original (.xlsx ou .xlsm)
        _, file_extension = os.path.splitext(uploaded_file.name)
        file_extension = file_extension.lower() if file_extension.lower() in [".xlsx", ".xlsm"] else ".xlsx"

        # 📝 Créer un fichier temporaire avec la bonne extension
        with tempfile.NamedTemporaryFile(delete=False, suffix=file_extension) as tmp_file:
            tmp_file.write(uploaded_file.read())
            local_path = tmp_file.name


            model = SentenceTransformer(MODEL_DIR_1)
            processor = TextProcessor(file_path_dictionnary)

            log_box = st.empty()

            def stream_log(msg):
                log_box.text(msg)

            try:
                start_time = time.time()
                progress.progress(20, text="🧠 Chargement du modèle et des données...")

                # 🔁 Traitement du fichier
                process_new_descriptions(
                    file_path_takalot_file=local_path,
                    embedding_model=model,
                    processor=processor,
                    reference_json=file_path_json,
                    log_fn=stream_log
                )

                progress.progress(90, text="💾 Sauvegarde du fichier modifié...")

                output_dir = "Data_Cabine/Output"
                os.makedirs(output_dir, exist_ok=True)
                final_path = generate_unique_filename(output_dir)
                os.makedirs(os.path.dirname(final_path), exist_ok=True)
                shutil.copy(local_path, final_path)

                progress.progress(100, text="✅ Terminé")
                elapsed = time.time() - start_time
                st.success(f"🎉 Traitement terminé en {elapsed:.2f} secondes")
                st.info(f"📁 Le fichier modifié a été sauvegardé ici : `{final_path}`")

            except Exception as e:
                st.error(f"❌ Erreur pendant le traitement : {e}")




































"""
import streamlit as st
import tempfile
import shutil
import time
import os

from sentence_transformers import SentenceTransformer
from text_processing import TextProcessor
from config import file_path_dictionnary, file_path_json, MODEL_DIR_1
from sentenceT import process_new_descriptions

from data_loader import CategoryEmbeddingBuilder
from config import file_path_categories

st.set_page_config(page_title="Classificateur Takalot", layout="centered")
st.title("📋 Classificateur de défauts de Cabine")

##########################################################################

if st.button("🧠 Générer un nouveau fichier d'embeddings"):
    with st.spinner("Génération des embeddings en cours..."):
        try:
            builder = CategoryEmbeddingBuilder(file_path_categories, MODEL_DIR_1)
            builder.save_outputs()
            st.success("✅ Embeddings créés avec succès et sauvegardés dans `embeddings.json`")
        except Exception as e:
            st.error(f"❌ Erreur lors de la génération des embeddings : {e}")

##########################################################################

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
        
         # ✅ Zone de log affichée dynamiquement
        log_box = st.empty()

        def stream_log(msg):
            log_box.text(msg)

        try:
            start_time = time.time()
            progress.progress(20, text="🧠 Chargement du modèle et des données...")

            process_new_descriptions(
                file_path_takalot_file=temp_input_path,
                embedding_model=model,
                processor=processor,
                reference_json=file_path_json,
                log_fn=stream_log  # 👈 Ajout du logger

            )

            progress.progress(90, text="💾 Sauvegarde du fichier modifié...")

            final_path = f"Data_Cabine/Output/fichier_modifie.xlsx"
            os.makedirs(os.path.dirname(final_path), exist_ok=True)
            shutil.copy(temp_input_path, final_path)

            progress.progress(100, text="✅ Terminé")
            elapsed = time.time() - start_time
            st.success(f"🎉 Traitement terminé en {elapsed:.2f} secondes")
            st.info(f"📁 Le fichier modifié a été sauvegardé ici : `{final_path}`")

        except Exception as e:
            st.error(f"❌ Erreur pendant le traitement : {e}")"""
