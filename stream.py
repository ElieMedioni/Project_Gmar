import streamlit as st
from datetime import datetime
import tempfile
import shutil
import time
import os

from sentence_transformers import SentenceTransformer
from text_processing import TextProcessor
from config import file_path_dictionnary, file_path_json, file_path_ata, MODEL_DIR_1
from sentenceT import process_new_descriptions

from data_loader import CategoryEmbeddingBuilder
from config import file_path_categories

st.set_page_config(page_title="סיווג תקלות קבינה", layout="centered")

st.markdown("""
    <style>
        .block-container {
            padding-top: 1rem !important;
        }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center; direction: rtl; font-family: \"Segoe UI\"; font-weight: bold;'>סיווג תקלות קבינה </h1><br>", unsafe_allow_html=True)

##########################################################################
#BOUTTON NOUVEAUX VECTEURS

col1, col2, col3 = st.columns([1, 1, 1])
with col3:
    if st.button("צור וקטורי ייחוס חדשים"):
        with col1:
            with st.spinner("...יצירת וקטורי ייחוס חדשים"):
                try:
                    builder = CategoryEmbeddingBuilder(file_path_categories, MODEL_DIR_1)
                    builder.save_outputs()
                    st.success("`embeddings.json`-וקטורים נוצרו ונשמרו בהצלחה ב")
                except Exception as e:
                    st.error(f"❌ {e} : שגיאה מסוג")
                
##########################################################################
#FONCTION NOM DU NOUVEAU DOSSIER

def generate_unique_filename(base_path, base_name="סיווג", ext=".xlsx"):
    # Obtenir la date du jour au format dd-mm-yyyy
    date_str = datetime.now().strftime("%d.%m.%Y")
    base_filename = f"{base_name}_{date_str}"

    i = 0
    while True:
        suffix = f"_{i}" if i > 0 else ""
        filename = f"{base_filename}{suffix}{ext}"
        full_path = os.path.join(base_path, filename)
        if not os.path.exists(full_path):
            return full_path
        i += 1

##########################################################################
#BOUTTON

st.markdown("<div style='text-align: right; direction: rtl; font-family: \"Segoe UI\"'>בחר קובץ אקסל כאן כדי לסווג אותו</div>", unsafe_allow_html=True)
uploaded_file = st.file_uploader(label="", type=["xlsx", "xlsm"])

if uploaded_file:
    st.success(f"{uploaded_file.name} : הקובץ התקבל בהצלחה")

    col1, col2, col3 = st.columns([1, 1, 1])
    with col3:
        run = st.button("🚀 לההריץ את המודל")

    if run:
        st.markdown("<div style='text-align: right; direction: rtl; font-family: \"Segoe UI\"; font-weight: bold;'>התחול של הסיווג</div>", unsafe_allow_html=True)
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
            
            try: 
                
                start_time = time.time()
                st.markdown("<div style='text-align: right; direction: rtl; font-family: \"Segoe UI\"; font-weight: bold;'>מריץ את המודל ומזהה תקלות</div>", unsafe_allow_html=True)
                
                log_box = st.empty()

                def stream_log(msg):
                    log_box.markdown(
                        f"<div style='text-align: right; direction: rtl; font-family: \"Segoe UI\";'>{msg}</div>",
                        unsafe_allow_html=True
                    )


                # 🔁 Traitement du fichier
                process_new_descriptions(
                    file_path_takalot_file=local_path,
                    embedding_model=model,
                    processor=processor,
                    ata=file_path_ata,
                    reference_json=file_path_json,
                    log_fn=stream_log
                )

                log_box = st.empty()

                st.markdown("<div style='text-align: right; direction: rtl; font-family: \"Segoe UI\";font-weight: bold;'>שמירת של התקלות מסווגות</div>", unsafe_allow_html=True)
                output_dir = "Data_Cabine/Output"
                os.makedirs(output_dir, exist_ok=True)
                final_path = generate_unique_filename(output_dir)
                os.makedirs(os.path.dirname(final_path), exist_ok=True)
                shutil.copy(local_path, final_path)
                
                st.markdown("<div style='text-align: right; direction: rtl; font-family: \"Segoe UI\";font-weight: bold;'>סיימנו✅ !</div>", unsafe_allow_html=True)
                st.markdown("*****")
                elapsed = time.time() - start_time
                
                st.markdown(f"""
                <div dir='rtl' style='
                    border: 2px solid #2ecc71;
                    padding: 12px;
                    border-radius: 8px;
                    font-family: "Segoe UI";
                    color: black;
                    font-weight: bold;
                '>
                🎉 הסיווג הושלם תוך {elapsed:.2f} שניות.<br>
                📁 הקובץ נשמר בהצלחה בשם: <b>{os.path.basename(final_path)}</b>
                </div>
                """, unsafe_allow_html=True)

            except Exception as e:
                st.error(f"❌ {e} : שגיאה מסוג")
