import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import random
from fpdf import FPDF
import plotly.io as pio
import os
import io

st.set_page_config(layout='wide')

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

# Cache pour charger les données Excel
@st.cache_data
def load_excel(file):
    return pd.read_excel(file)

# Cache pour préparer et nettoyer les données
@st.cache_data
def prepare_data(df):
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

    return df

# Cache pour calculer les couleurs par compagnie
@st.cache_data
def assign_colors(compagnies):
    compagnie_colors = {}
    for compagnie in compagnies:
        if compagnie not in compagnie_colors:
            compagnie_colors[compagnie] = generate_random_color()
    return compagnie_colors


def export_gantt_to_pdf_one_page_streamlit(fig_global, fig_charge, fig_escabeaux):
    """
    Exporte les trois diagrammes de Gantt sur une seule page PDF, et retourne un objet binaire pour Streamlit.

    :param fig_global: Graphique de charge globale (Plotly Figure).
    :param fig_charge: Graphique de charge par compagnie (Plotly Figure).
    :param fig_escabeaux: Graphique des besoins en escabeaux (Plotly Figure).
    :return: BytesIO object contenant le PDF.
    """
    # Créer un répertoire temporaire pour enregistrer les images
    temp_dir = "temp_gantt_images"
    os.makedirs(temp_dir, exist_ok=True)

    # Sauvegarder les figures en tant qu'images
    fig_global_path = os.path.join(temp_dir, "charge_globale.png")
    fig_charge_path = os.path.join(temp_dir, "charge_par_compagnie.png")
    fig_escabeaux_path = os.path.join(temp_dir, "besoins_escabeaux.png")

    # Exporter les graphiques Plotly en fichiers PNG
    pio.write_image(fig_global, fig_global_path, format="png", width=2000, height=400)
    pio.write_image(fig_charge, fig_charge_path, format="png", width=2000, height=400)
    pio.write_image(fig_escabeaux, fig_escabeaux_path, format="png", width=2000, height=400)

    # Initialiser le PDF
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(0, 10, f"Diagrammes de Gantt pour la journée du {selected_date}", ln=True, align='C')

    # Ajouter les images au PDF (organisées verticalement)
    y_position = 20
    for img_path in [fig_global_path, fig_charge_path, fig_escabeaux_path]:
        pdf.image(img_path, x=10, y=y_position, w=270)  # Ajuste largeur pour A4 paysage
        y_position += 70  # Espacement vertical entre les graphiques

    # Obtenir le contenu PDF en tant que bytes
    pdf_data = pdf.output(dest='S').encode('latin1')  # Générer un fichier PDF dans un buffer en mémoire

    # Supprimer les images temporaires
    for file in [fig_global_path, fig_charge_path, fig_escabeaux_path]:
        os.remove(file)
    os.rmdir(temp_dir)

    # Convertir les données en objet BytesIO
    pdf_buffer = io.BytesIO(pdf_data)

    return pdf_buffer

# Streamlit App
st.title("Analyse des Vols avec Courbes de Charge, Gantt et Statistiques")

# Chargement du fichier Excel
uploaded_file = st.sidebar.file_uploader("Importez le fichier des vols", type=["xlsx", "xls"])

if uploaded_file:
    # Charger les données Excel
    df = load_excel(uploaded_file)

    # Préparer les données
    df = prepare_data(df)

    # Génération des couleurs par compagnie
    compagnies = df["Compagnie"].dropna().unique()
    compagnie_colors = assign_colors(compagnies)
    df["Color"] = df["Compagnie"].map(compagnie_colors)

    # Ajout d'un filtre pour choisir une date spécifique
    unique_dates = df["Date"].dt.date.unique()
    selected_date = st.sidebar.selectbox("Choisissez une date :", unique_dates)

    if selected_date:
        # Filtrer les données pour la date sélectionnée
        df_filtered = df[df["Date"].dt.date == selected_date]

        # **Courbe de charge globale**
        st.subheader(f"Courbe de charge globale pour le {selected_date}")
        time_range = pd.date_range(f"{selected_date} 05:00", f"{selected_date} 23:59", freq="1min")
        charge_global = []
        for time in time_range:
            count = ((df_filtered["Start"] <= time) & (df_filtered["End"] >= time)).sum()
            charge_global.append(count)

        # Création de la courbe de charge globale
        fig_global = go.Figure()
        fig_global.add_trace(
            go.Scatter(
                x=time_range,
                y=charge_global,
                mode="lines",
                name="Charge Globale",
                line=dict(color="blue"),
            )
        )
        fig_global.update_layout(
            title="Courbe de Charge Globale des Vols",
            xaxis_title="Heures de la journée",
            yaxis_title="Nombre de vols simultanés",
            xaxis=dict(tickformat="%H:%M"),
            yaxis=dict(tickformat="d"),
            width=1300,
            height=500,
        )
        st.plotly_chart(fig_global)

        # **Courbes de charge multi-tracés par compagnie**
        st.subheader(f"Courbes de charge par compagnie pour le {selected_date}")
        fig_charge = go.Figure()

        # Pour chaque compagnie, tracer sa courbe
        for compagnie in compagnies:
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
            width=1400,
            height=500,
        )
        st.plotly_chart(fig_charge)

        # **Besoins en escabeaux (Gantt et données)**
        st.subheader(f"Besoins en escabeaux pour le {selected_date}")
        time_range_escabeaux = pd.date_range(f"{selected_date} 05:00", f"{selected_date} 23:59", freq="15min")
        besoin_escabeaux = []
        for time in time_range_escabeaux:
            active_vols = df_filtered[(df_filtered["Start"] <= time) & (df_filtered["End"] >= time)]
            count_fr = (active_vols["Compagnie"] == "FR").sum()  # 1 escabeau pour "FR"
            count_other = (active_vols["Compagnie"] != "FR").sum()  # 2 escabeaux pour les autres
            total_escabeaux = count_fr * 1 + count_other * 2
            besoin_escabeaux.append(total_escabeaux)

        # DataFrame des besoins en escabeaux
        escabeaux_df = pd.DataFrame({
            "Heure": time_range_escabeaux,
            "Besoins en Escabeaux": besoin_escabeaux
        })

        # Gantt pour les escabeaux
        fig_escabeaux = go.Figure()
        fig_escabeaux.add_trace(
            go.Bar(
                x=escabeaux_df["Heure"],
                y=escabeaux_df["Besoins en Escabeaux"],
                name="Besoins en Escabeaux",
                marker=dict(color="darkgreen"),
            )
        )
        fig_escabeaux.update_layout(
            title="Gantt des Besoins en Escabeaux",
            xaxis_title="Heures de la journée",
            yaxis_title="Nombre d'escabeaux nécessaires",
            xaxis=dict(tickformat="%H:%M"),
            yaxis=dict(tickformat="d"),
            width=1300,
            height=500,
        )
        st.plotly_chart(fig_escabeaux)

        # Bouton pour télécharger le PDF
        if st.sidebar.button("Générer et Télécharger le PDF des Graphs"):
            with st.spinner("Génération du PDF en cours..."):
                pdf_buffer = export_gantt_to_pdf_one_page_streamlit(fig_global, fig_charge, fig_escabeaux)

                # Ajouter un lien de téléchargement
                st.sidebar.download_button(
                    label="Télécharger le PDF",
                    data=pdf_buffer,
                    file_name="gantt_charts.pdf",
                    mime="application/pdf",
                )

        st.subheader(f"Tableaux des donnée pour le {selected_date}")
        if st.checkbox("Afficher les tableaux :"):
            # Affichage du DataFrame des besoins en escabeaux
            st.subheader("Détails des Besoins en Escabeaux par Tranches de 15 Minutes")
            st.dataframe(escabeaux_df)

            # Affichage des données filtrées
            st.subheader("Données de la journée sélectionnée")
            st.dataframe(df_filtered)

        # **Statistiques de la journée**
        st.subheader(f"Statistiques pour le {selected_date}")
        nb_arrivees = df_filtered["Arr"].notna().sum()
        nb_departs = df_filtered["Dép"].notna().sum()
        st.markdown(f"**Nombre d'arrivées :** {nb_arrivees}")
        st.markdown(f"**Nombre de départs :** {nb_departs}")

        # Ventilation par compagnie
        compagnie_counts = df_filtered["Compagnie"].value_counts()
        fig_compagnie = go.Figure(
            data=[go.Pie(labels=compagnie_counts.index, values=compagnie_counts.values, hole=0.4)]
        )
        fig_compagnie.update_layout(title="Répartition des vols par compagnie")
        st.plotly_chart(fig_compagnie)
