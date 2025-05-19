"""import os
from sentence_transformers import SentenceTransformer
from config import MODEL_DIR_1, MODEL_DIR_2

# Vérifier si le dossier existe, sinon le créer
os.makedirs(MODEL_DIR_1, exist_ok=True)

# Télécharger et sauvegarder le modèle dans le dossier défini
model = SentenceTransformer("sentence-transformers/all-mpnet-base-v2", cache_folder=MODEL_DIR_1)

print(f"Le modèle est enregistré dans : {MODEL_DIR_2}")"""

import os
from sentence_transformers import SentenceTransformer

# Dictionnaire des modèles à télécharger
models = {
    "all-mpnet-base-v2": "sentence-transformers/all-mpnet-base-v2",
   # "all-MiniLM-L6-v2": "sentence-transformers/paraphrase-MiniLM-L6-v2"
}

for local_dir, model_name in models.items():
    print(f"Téléchargement de {model_name} vers {local_dir}...")

    # Crée le dossier si besoin
    os.makedirs(local_dir, exist_ok=True)

    # Télécharge et sauvegarde dans le dossier local
    model = SentenceTransformer(model_name)
    model.save(local_dir)

    print(f"✅ {model_name} sauvegardé avec succès dans ./{local_dir}\n")
