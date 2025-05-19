from flask import Flask, request, jsonify
from sentence_transformers import SentenceTransformer
from config import MODEL_DIR

# Chargement direct du modèle depuis Hugging Face
model = SentenceTransformer(MODEL_DIR)

# Création de l'application Flask
app = Flask(__name__)

@app.route("/embed", methods=["POST"])
def embed_text():
    try:
        data = request.get_json()
        if "text" not in data:
            return jsonify({"error": "Le champ 'text' est requis."}), 400
        
        text = data["text"]
        if not isinstance(text, str):
            return jsonify({"error": "Le champ 'text' doit être une chaîne de caractères."}), 400
        # Génération de l'embedding
        embedding = model.encode(text).tolist()
        return jsonify({"text": text, "embedding": embedding})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)  # Désactiver le mode debug