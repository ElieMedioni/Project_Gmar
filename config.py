import os

# Chemin vers le dossier où les modèles sont sauvegardés
MODEL_DIR_1 = "Models/all-mpnet-base-v2"

# on peut changer "data_cabine" avec le dossier désiré, et que les fichiers ait le meme format pour que ca s'adapte
file_path_categories = os.path.join("Data_Cabine", "Categories.xlsx")
file_path_ata = os.path.join("Data_Cabine", "Ata Chapter.xlsx")
file_path_dictionnary = os.path.join("Data_Cabine", "Dictionnary.xlsx")
file_path_json = os.path.join("Data_Cabine", "embeddings.json")
