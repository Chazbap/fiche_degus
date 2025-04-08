import pandas as pd
import streamlit as st

# === Chemins des fichiers ===
file_path_cuves = "C:/Users/bchazeaud/OneDrive - vrankenpommery.fr/Pinglestone Estate/Python/2711.xlsx"
file_path_codes = "C:/Users/bchazeaud/OneDrive - vrankenpommery.fr/Pinglestone Estate/Python/code_produit.xlsx"

# === Chargement des fichiers ===
df_cuves = pd.read_excel(file_path_cuves)
df_codes = pd.read_excel(file_path_codes)

# Fusionner les données de cuves et codes produits
df_cuves = df_cuves.merge(df_codes[['Code Produit en Cuve', 'Clé Produit en Cuve']], left_on='Produit', right_on='Code Produit en Cuve', how='left')

# === Streamlit ===
st.title("📝 Générateur de Fiches de Dégustation")

# Demander le nombre de dégustateurs
st.subheader("1️⃣ Nombre de dégustateurs")

num_degustateurs = st.number_input("Entrez le nombre de dégustateurs", min_value=1, max_value=10, value=1, step=1)

if num_degustateurs < 1:
    st.warning("Veuillez entrer un nombre valide de dégustateurs.")
    st.stop()

# === Collecte des informations des dégustateurs ===
degustateurs = []
for i in range(num_degustateurs):
    degustateur = st.text_input(f"Entrez le prénom ou les initiales du dégustateur #{i+1}")
    degustateurs.append(degustateur)

# Si un dégustateur n'a pas été ajouté, afficher un message d'avertissement
if any([degustateur.strip() == "" for degustateur in degustateurs]):
    st.warning("Tous les dégustateurs doivent être identifiés avant de continuer.")
    st.stop()

# Sélection des cuves
st.subheader("2️⃣ Sélection des cuves")

cuves_disponibles = df_cuves['N° Cuve'].dropna().unique().tolist()

cuves_selectionnees = st.multiselect(
    "Choisissez les cuves à déguster :",
    options=cuves_disponibles
)

if not cuves_selectionnees:
    st.warning("Veuillez sélectionner au moins une cuve pour continuer.")
    st.stop()

# Initialisation des notes des dégustateurs
if 'notes' not in st.session_state:
    st.session_state.notes = {degustateur: {} for degustateur in degustateurs}

# === Collecte des notes et affichage des moyennes ===
st.subheader("3️⃣ Saisie des notes et affichage des moyennes")

# Dégustateur actuel
degustateur_courant = st.selectbox("Sélectionner votre nom", options=degustateurs)

# Affichage des curseurs pour chaque cuve
for cuve in cuves_selectionnees:
    # Obtenir les infos associées à chaque cuve
    cuve_data = df_cuves[df_cuves['N° Cuve'] == cuve].iloc[0]
    code_produit = cuve_data['Code Produit en Cuve']
    volume = cuve_data['En Stock']

    st.write(f"Cuve: {cuve} | Code produit: {code_produit} | Volume: {volume} L")

    # Collecte des notes
    notes = {}
    notes['Tension'] = st.slider(f"Tension pour {cuve} (0-5)", 0, 5, 0)
    notes['Volume'] = st.slider(f"Volume pour {cuve} (0-5)", 0, 5, 0)
    notes['Amertume'] = st.slider(f"Amertume pour {cuve} (0-5)", 0, 5, 0)
    notes['Finesse'] = st.slider(f"Finesse pour {cuve} (0-5)", 0, 5, 0)
    notes['Défaut'] = st.slider(f"Défaut pour {cuve} (0-1)", 0, 1, 0)
    notes['Note Globale'] = st.slider(f"Note Globale pour {cuve} (0-10)", 0, 10, 0)

    # Enregistrer les notes dans session_state pour ce dégustateur et cuve
    st.session_state.notes[degustateur_courant][cuve] = notes

    # Calcul des moyennes des notes de tous les dégustateurs pour chaque catégorie
    moyenne_notes = {key: [] for key in notes.keys()}

    for degustateur in degustateurs:
        if cuve in st.session_state.notes[degustateur]:
            for key in notes.keys():
                moyenne_notes[key].append(st.session_state.notes[degustateur][cuve][key])

    moyenne_notes = {key: sum(values)/len(values) if values else 0 for key, values in moyenne_notes.items()}

    # Affichage des moyennes des autres dégustateurs
    col1, col2 = st.columns([3, 1])  # Deux colonnes : une plus large pour les curseurs et une plus étroite pour la moyenne
    with col1:
        st.write(f"### Notes de {degustateur_courant} pour la Cuve {cuve}")
    with col2:
        st.write("### Moyenne des notes de tous les dégustateurs")

    with col1:
        st.write(f"Tension: {notes['Tension']} | Volume: {notes['Volume']} | Amertume: {notes['Amertume']} | Finesse: {notes['Finesse']} | Défaut: {notes['Défaut']} | Note Globale: {notes['Note Globale']}")
    
    with col2:
        st.write(f"Tension: {moyenne_notes['Tension']:.2f}")
        st.write(f"Volume: {moyenne_notes['Volume']:.2f}")
        st.write(f"Amertume: {moyenne_notes['Amertume']:.2f}")
        st.write(f"Finesse: {moyenne_notes['Finesse']:.2f}")
        st.write(f"Défaut: {moyenne_notes['Défaut']:.2f}")
        st.write(f"Note Globale: {moyenne_notes['Note Globale']:.2f}")

# Téléchargement des résultats (si souhaité)
if st.button("Télécharger les résultats sous forme de fichier Excel"):
    # Créer un dataframe avec les résultats
    data_rows = []
    for cuve in cuves_selectionnees:
        for degustateur, notes in st.session_state.notes.items():
            if cuve in notes:
                row = [cuve, degustateur] + list(notes[cuve].values())
                data_rows.append(row)

    df_resultats = pd.DataFrame(data_rows, columns=['Cuve', 'Dégustateur', 'Tension', 'Volume', 'Amertume', 'Finesse', 'Défaut', 'Note Globale'])
    fichier_excel = "resultats_degustation.xlsx"
    df_resultats.to_excel(fichier_excel, index=False)

    with open(fichier_excel, "rb") as f:
        st.download_button(
            label="Télécharger les résultats",
            data=f,
            file_name=fichier_excel,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
