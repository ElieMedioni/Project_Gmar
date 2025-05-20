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

from config import MODEL_DIR_1

os.makedirs(MODEL_DIR_1, exist_ok=True)

model = SentenceTransformer("sentence-transformers/all-mpnet-base-v2", cache_folder=MODEL_DIR_1)

print(f"Le modèle est enregistré dans : {MODEL_DIR_1}")