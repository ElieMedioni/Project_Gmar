import pandas as pd
import numpy as np
import json
import time
from tqdm import tqdm
from scipy.spatial.distance import cdist
from sentence_transformers import SentenceTransformer
from config import file_path_dictionnary, file_path_json, file_path_json_2, file_path_takalot_file, file_path_output_file, MODEL_DIR_1, MODEL_DIR_2
from text_processing import TextProcessor


def load_embeddings_from_json(json_path):
    with open(json_path, "r", encoding="utf-8") as f:
        return json.load(f)

def process_new_descriptions(file_path_takalot_file, file_path_output_file, embedding_model, processor, reference_json):
    print("Chargement des données...")
    df = pd.read_excel(file_path_takalot_file, header=None, names=["RAW_DESCRIPTION"])
    reference_embeddings = load_embeddings_from_json(reference_json)

    print("Nettoyage des descriptions...")
    df["CLEANED_DESCRIPTION"] = df["RAW_DESCRIPTION"].astype(str).apply(processor.clean_text)

    print("Génération des embeddings...")
    takalot_embeddings = embedding_model.encode(df["CLEANED_DESCRIPTION"].tolist(), normalize_embeddings=True, show_progress_bar=True, batch_size=32)

    ref_keys = list(reference_embeddings.keys())
    ref_matrix = np.array([reference_embeddings[k] for k in ref_keys])

    """print("Calcul des distances (cosine)...")
    dist_matrix = cdist(takalot_embeddings, ref_matrix, metric="cosine")"""
    
    print("Calcul des distances (euclidean)...")
    dist_matrix = cdist(takalot_embeddings, ref_matrix, metric="euclidean")

    print("Extraction des meilleurs correspondances...")
    top_2_idx = np.argsort(dist_matrix, axis=1)[:, :2]
    #top_2_idx = np.argsort(dist_matrix, axis=1)[:, :1]
    top_2_dists = np.take_along_axis(dist_matrix, top_2_idx, axis=1)
    
    results = []
    for i in tqdm(range(len(df)), desc="Traitement des lignes"):

        best_1 = ref_keys[top_2_idx[i, 0]]
        best_2 = ref_keys[top_2_idx[i, 1]]
        dist_1 = round(top_2_dists[i, 0], 4)
        dist_2 = round(top_2_dists[i, 1], 4)

        cat_1, subcat_1 = best_1.split(" | ")
        cat_2, subcat_2 = best_2.split(" | ")

        results.append({
            "Description": df["RAW_DESCRIPTION"].iloc[i],
            "Cleaned Description": df["CLEANED_DESCRIPTION"].iloc[i],
            "1 Main category": cat_1,
            "1 Sub Category": subcat_1,
            "1 Distance": dist_1,
            "2 Main category": cat_2,
            "2 Sub Category": subcat_2,
            "2 Distance": dist_2,
        })


    pd.DataFrame(results).to_excel(file_path_output_file, index=False)
    print(f"\nRésultat enregistré dans : {file_path_output_file}")


if __name__ == "__main__":
    try:
        start_time = time.time()
        
        # MODEL_DIR_1 = "all-mpnet-base-v2"
        model = SentenceTransformer(MODEL_DIR_1)
        processor = TextProcessor(file_path_dictionnary)

        process_new_descriptions(
            file_path_takalot_file,
            file_path_output_file,
            embedding_model=model,
            processor=processor,
            reference_json=file_path_json
        )
        
        # MODEL_DIR_2 = "all-MiniLM-L6-v2"
        """model = SentenceTransformer(MODEL_DIR_2)
        processor = TextProcessor(file_path_dictionnary)

        process_new_descriptions(
            file_path_takalot_file,
            file_path_output_file,
            embedding_model=model,
            processor=processor,
            reference_json=file_path_json_2
        )"""

        print(f"Temps total : {time.time() - start_time:.2f} secondes")

    except Exception as e:
        print(f"❌ Erreur : {e}")
