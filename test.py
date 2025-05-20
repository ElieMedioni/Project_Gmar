import xlwings as xw

# 🗂️ Chemin vers ton fichier Excel à tester
path = "Data_Cabine/Row_Data_2025.xlsm"

try:
    app = xw.App(visible=True)
    wb = app.books.open(path)
    ws = wb.sheets[0]  # 👉 Modifie la première feuille (index 0)

    # 📝 Test : écrire "Hello from xlwings!" dans la cellule A1
    ws.range("A1").value = "Hello from xlwings!"

    # 💾 Sauvegarde le fichier
    wb.save()
    print("✅ Écriture réussie. Vérifie la cellule A1 du fichier.")

except Exception as e:
    print("❌ Erreur pendant l’écriture :", e)
