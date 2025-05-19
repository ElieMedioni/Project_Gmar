import streamlit as st

# Interface utilisateur
st.title("Mon Application Python")

# Entrée utilisateur
input_value = st.text_input("Entrez une valeur:")

# Bouton d'exécution
if st.button("Exécuter le code backend"):
    # Remplacer ce bloc par ton code backend
    result = f"Vous avez entré: {input_value}"
    
    # Affichage du résultat
    st.success(result)
