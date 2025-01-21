import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import random

# Fonction pour convertir les heures au format hh:mm
def format_hhmm_to_hhmm(time):
    if pd.isna(time) or time == "-":
        return None
    time = str(time).zfill(4)  # S'assurer d'avoir 4 chiffres
    return f"{time[:2]}:{time[2:]}"

# Fonction pour calculer les colonnes Start et End selon les conditions spécifiées
def compute_start_end(row):
    start = row["Start"]
    end = row["End"]

    if pd.isna(row["Arr"]) and not pd.isna(row["Dép"]):
        # Si Arr est None, Start = Dép - 35 minutes
        start = row["End"] - pd.Timedelta(minutes=35)
    if pd.isna(row["Dép"]) and not pd.isna(row["Arr"]):
        # Si Dép est None, End = Arr + 35 minutes
        end = row["Start"] + pd.Timedelta(minutes=35)

    return start, end

# Extraire l'indicatif de la compagnie
def extract_company_code(flight_number):
    if isinstance(flight_number, str) and flight_number != "-":
        return "".join([char for char in flight_number if char.isalpha()])
    return None

# Générer une couleur aléatoire
def generate_random_color():
    return f"#{random.randint(0, 0xFFFFFF):06x}"

# Dictionnaire fixe des couleurs des compagnies (modifiable si nécessaire)
compagnie_colors = {
    "FR": "#FF5733",  # Exemple : Couleur fixe pour la compagnie "FR"
    "W6": "#33FF57",  # Exemple : Couleur fixe pour la compagnie "W6"
    "LH": "#3357FF",  # Exemple : Couleur fixe pour la compagnie "LH"
}

# Streamlit App
st.title("Analyse des Vols avec Courbes de Charge et Statistiques")

# Chargement du fichier Excel
uploaded_file = st.file_uploader("Importez un fichier Excel contenant les données des vols", type=["xlsx", "xls"])

if uploaded_file:
    # Lecture du fichier Excel
    df = pd.read_excel(uploaded_file)

    # Nettoyage et préparation des données
    df["Arr"] = df["Arr"].apply(format_hhmm_to_hhmm)
    df["Dép"] = df["Dép"].apply(format_hhmm_to_hhmm)

    # Générer les colonnes Start et End initialement
    df["Start"] = pd.to_datetime(df["Date"].dt.strftime('%Y-%m-%d') + " " + df["Arr"].fillna("00:00"), errors="coerce")
    df["End"] = pd.to_datetime(df["Date"].dt.strftime('%Y-%m-%d') + " " + df["Dép"].fillna("00:00"), errors="coerce")

    # Appliquer les ajustements pour les valeurs manquantes
    df[["Start", "End"]] = df.apply(lambda row: compute_start_end(row), axis=1, result_type="expand")

    # Extraire le code de la compagnie aérienne
    df["Compagnie"] = df.apply(
        lambda row: extract_company_code(row["n°arr"])
        if pd.notna(row["n°arr"]) and row["n°arr"] != "-"
        else extract_company_code(row["n°dep"])
        if pd.notna(row["n°dep"]) and row["n°dep"] != "-"
        else None,
        axis=1
    )

    # Supprimer les lignes où les intervalles ne peuvent pas être créés
    df = df.dropna(subset=["Start", "End"])

    # Assigner une couleur fixe ou générer une nouvelle couleur pour les compagnies absentes du dictionnaire
    unique_compagnies = df["Compagnie"].dropna().unique()
    for compagnie in unique_compagnies:
        if compagnie not in compagnie_colors:
            compagnie_colors[compagnie] = generate_random_color()

    # Assigner les couleurs aux données
    df["Color"] = df["Compagnie"].map(compagnie_colors)

    # Ajout d'un filtre pour choisir une date spécifique
    unique_dates = df["Date"].dt.date.unique()
    selected_date = st.selectbox("Choisissez une date :", unique_dates)

    if selected_date:
        # Filtrer les données pour la date sélectionnée
        df_filtered = df[df["Date"].dt.date == selected_date]

        # **Courbes de charge multi-tracés par compagnie**
        st.subheader(f"Courbes de charge par compagnie pour le {selected_date}")
        fig_charge = go.Figure()

        # Pour chaque compagnie, tracer sa courbe
        time_range = pd.date_range(f"{selected_date} 00:00", f"{selected_date} 23:59", freq="1min")
        for compagnie in unique_compagnies:
            df_compagnie = df_filtered[df_filtered["Compagnie"] == compagnie]
            charge = []
            for time in time_range:
                count = ((df_compagnie["Start"] <= time) & (df_compagnie["End"] >= time)).sum()
                charge.append(count)

            fig_charge.add_trace(
                go.Scatter(
                    x=time_range,
                    y=charge,
                    mode="lines",
                    name=compagnie,
                    line=dict(color=compagnie_colors[compagnie]),
                )
            )

        fig_charge.update_layout(
            title="Courbes de Charge des Vols par Compagnie",
            xaxis_title="Heures de la journée",
            yaxis_title="Nombre de vols simultanés",
            xaxis=dict(tickformat="%H:%M"),
            yaxis=dict(tickformat="d"),
            legend_title="Compagnies",
        )
        st.plotly_chart(fig_charge)

        # Affichage des couleurs fixes des compagnies
        st.subheader("Couleurs des Compagnies (fixes)")
        st.dataframe(pd.DataFrame(list(compagnie_colors.items()), columns=["Compagnie", "Couleur"]))

        # Le reste du code reste inchangé (statistiques, courbes de charge, besoins en escabeaux, etc.)
