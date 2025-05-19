import requests
import json

# URL de l'API
API_URL = "http://localhost:5000/embed"

def get_embedding(text):
    """Envoie un texte à l'API et récupère son embedding."""
    payload = {"text": text}
    headers = {"Content-Type": "application/json"}

    try:
        response = requests.post(API_URL
                                 , data=json.dumps(payload)
                                 , headers=headers)
        response.raise_for_status()  # Lève une exception si le statut HTTP est une erreur
        return response.json()  # Retourne les embeddings
    except requests.exceptions.RequestException as e:
        print("Erreur lors de la requête :", e)
        return None

base_text = "ca va frero?"
result = get_embedding(base_text)

print(result)