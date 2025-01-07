import pandas as pd
import streamlit as st
from datetime import datetime, timedelta, time
import plotly.express as px
import streamlit.components.v1 as components
import io
import plotly.graph_objects as go

st.set_page_config(layout='wide')

@st.cache_data
def cached_preprocess_data(data):
    """Fonction mise en cache pour prétraiter les données."""
    return preprocess_data(data)


def parse_time(time_value, date=None):
    """
    Parse les valeurs de temps en conservant la date spécifiée.

    Args:
        time_value: La valeur de temps à parser
        date: La date à utiliser (datetime.date)
    """
    if pd.isna(time_value):
        return None

    if date is None:
        date = datetime.today().date()

    if isinstance(time_value, str):
        try:
            parsed_time = datetime.strptime(time_value, '%H:%M:%S').time()
        except ValueError:
            try:
                parsed_time = datetime.strptime(time_value, '%H:%M').time()
            except ValueError:
                return None
        return datetime.combine(date, parsed_time)
    elif isinstance(time_value, datetime):
        return datetime.combine(date, time_value.time())
    elif isinstance(time_value, timedelta):
        base_time = (datetime.min + time_value).time()
        return datetime.combine(date, base_time)
    elif isinstance(time_value, time):
        return datetime.combine(date, time_value)
    return None

def format_date(date_str):
    date_obj = pd.to_datetime(date_str)
    return date_obj.strftime('%d/%m/%Y')


def preprocess_data(data):
    df = data.copy()

    # Convertir la colonne DATE en datetime
    df['DATE'] = pd.to_datetime(df['DATE'])

    # Utiliser la date correspondante pour chaque vol
    df['HA'] = df.apply(lambda row: parse_time(row['HA'], row['DATE'].date()), axis=1)
    df['HD'] = df.apply(lambda row: parse_time(row['HD'], row['DATE'].date()), axis=1)

    night_stop_mask = df['HA'].notna() & df['HD'].isna()
    df.loc[night_stop_mask, 'HD'] = df.loc[night_stop_mask, 'HA'].apply(
        lambda x: x + timedelta(minutes=30)
    )

    depart_sec_mask = df['HD'].notna() & df['HA'].isna()
    df.loc[depart_sec_mask, 'HA'] = df.loc[depart_sec_mask, 'HD'].apply(
        lambda x: x - timedelta(minutes=30)
    )

    df['Duration'] = (df['HD'] - df['HA']).dt.total_seconds() / 60
    df['Company'] = df['VOLD'].str.extract(r'^([A-Za-z]+)', expand=False)
    night_stop_company = df['VOLA'].str.extract(r'^([A-Za-z]+)', expand=False)
    df.loc[df['Company'].isna(), 'Company'] = night_stop_company

    return df[df['HA'].notna() | df['HD'].notna()]

def calculate_flight_stats(data):
    flights_per_company = data.groupby(['Company']).size().reset_index(name='Nombre de vols')
    total_flights = flights_per_company['Nombre de vols'].sum()
    if 'PAX' in data.columns:
        pax_per_company = data.groupby(['Company'])['PAX'].sum().reset_index(name='Nombre de passagers')
        total_pax = pax_per_company['Nombre de passagers'].sum()
    else:
        pax_per_company = pd.DataFrame(columns=['Company', 'Nombre de passagers'])
        total_pax = 0
    return flights_per_company, pax_per_company, total_pax, total_flights

def create_echarts_html(data, title, type_data):
    chart_data = data.to_dict('records')
    html = f"""
    <div id="chart" style="width:100%; height:400px;"></div>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/echarts/5.4.3/echarts.min.js"></script>
    <script>
    var chart = echarts.init(document.getElementById('chart'));
    var option = {{
        title: {{ text: '{title}', left: 'center' }},
        tooltip: {{ trigger: 'item', formatter: '{{b}}: {{c}} {type_data} ({{d}}%)' }},
        legend: {{ orient: 'vertical', left: 'left' }},
        series: [
            {{
                type: 'pie',
                radius: '50%',
                data: {[{'value': d[list(d.keys())[1]], 'name': d[list(d.keys())[0]]} for d in chart_data]},
                emphasis: {{
                    itemStyle: {{
                        shadowBlur: 10,
                        shadowOffsetX: 0,
                        shadowColor: 'rgba(0, 0, 0, 0.5)'
                    }}
                }}
            }}
        ]
    }};
    chart.setOption(option);
    window.addEventListener('resize', function() {{
        chart.resize();
    }});
    </script>
    """
    return html


def create_interactive_gantt(data, selected_date):
    data['HA_str'] = data['HA'].dt.strftime('%H:%M:%S')
    data['HD_str'] = data['HD'].dt.strftime('%H:%M:%S')
    data['Flight_Type'] = 'Normal'
    data.loc[pd.isna(data['VOLA']) & pd.notna(data['VOLD']), 'Flight_Type'] = 'Depart_Sec'
    data.loc[pd.notna(data['VOLA']) & pd.isna(data['VOLD']), 'Flight_Type'] = 'Night_Stop'

    def annotation(row):
        if pd.notna(row['VOLD']) and pd.notna(row['VOLA']):
            return f"{row['VOLD']}"
        elif pd.isna(row['VOLD']):
            return f"{row['VOLA']}"
        else:
            return f"{row['VOLD']}"

    data['Annotation'] = data.apply(annotation, axis=1)
    data['Vacation Line'] = 0

    for idx, row in data.iterrows():
        for line in range(len(data['Vacation Line'].unique()) + 1):
            overlapping = data[(data['Vacation Line'] == line) & ((data['HA'] < row['HD']) & (data['HD'] > row['HA']))]
            if overlapping.empty:
                data.at[idx, 'Vacation Line'] = line
                break

    # Définition d'une palette de couleurs fixe pour les compagnies
    colors = [
        '#FF9999', '#66B2FF', '#99FF99', '#FFCC99', '#FF99CC',
        '#99CCFF', '#FFB366', '#FF99FF', '#99FFCC', '#FFB3B3'
    ]
    company_colors = {
        company: colors[i % len(colors)]
        for i, company in enumerate(data['Company'].unique())
    }

    # Création du timeline avec la palette de couleurs personnalisée
    fig = px.timeline(
        data,
        x_start="HA",
        x_end="HD",
        y="Vacation Line",
        color="Company",
        color_discrete_map=company_colors,
        width=1500,
        height=600,
        title=f"Planning des vols du {selected_date.strftime('%d/%m/%Y')}",
        labels={"Company": "Compagnie", "Vacation Line": "Ligne de vacation"},
        hover_data={"HA_str": True, "HD_str": True, "VOLD": True, "VOLA": True, "DEST": True, "ORG": True,
                    "Flight_Type": True}
    )

    # Personnalisation de l'apparence des barres
    fig.update_traces(
        marker_line_color='rgb(8,48,107)',
        marker_line_width=1.5,
        opacity=0.85
    )

    # Ajout des annotations de vol et des liserets
    for i, row in data.iterrows():
        # Annotation du numéro de vol
        fig.add_annotation(
            x=row['HA'] + (row['HD'] - row['HA']) / 2,
            y=row['Vacation Line'],
            text=row['Annotation'],
            showarrow=False,
            font=dict(size=10, color="black"),
            align="center",
            xanchor="center",
            yanchor="middle"
        )

        # Liserets pour les vols spéciaux
        if row['Flight_Type'] == 'Depart_Sec':
            fig.add_shape(
                type="line",
                x0=row['HD'],
                x1=row['HD'],
                y0=row['Vacation Line'] - 0.4,
                y1=row['Vacation Line'] + 0.4,
                line=dict(color="#FF0000", width=4),
                layer="above"
            )
        elif row['Flight_Type'] == 'Night_Stop':
            fig.add_shape(
                type="line",
                x0=row['HA'],
                x1=row['HA'],
                y0=row['Vacation Line'] - 0.4,
                y1=row['Vacation Line'] + 0.4,
                line=dict(color="#334cff", width=4),
                layer="above"
            )

    # Mise à jour de la mise en page
    fig.update_yaxes(
        title_text="Lignes de vacation",
        categoryorder="total ascending",
        gridcolor='rgba(128, 128, 128, 0.2)'
    )

    fig.update_xaxes(
        gridcolor='rgba(128, 128, 128, 0.2)',
        tickformat='%H:%M'
    )

    fig.update_layout(
        plot_bgcolor='rgba(240, 248, 255, 0.8)',
        paper_bgcolor='white',
        margin=dict(l=50, r=50, t=50, b=50),
        showlegend=True,
        legend=dict(
            bgcolor='rgba(255, 255, 255, 0.9)',
            bordercolor='gray',
            borderwidth=1
        ),
        hoverlabel=dict(
            bgcolor="white",
            font_size=12
        )
    )

    # Légende des types de vols
    fig.add_annotation(
        x=1.02,
        y=1,
        xref="paper",
        yref="paper",
        text="Types de vols:<br>Liseret Rouge à droite = Départ Sec<br>Liseret bleu à gauche= Night Stop",
        showarrow=False,
        font=dict(size=12),
        align="left"
    )

    return fig


def export_gantt_to_pdf(fig):
    # Créer une copie de la figure pour l'export
    export_fig = fig

    # Mettre à jour les paramètres spécifiques pour l'export PDF
    export_fig.update_layout(
        paper_bgcolor='white',
        plot_bgcolor='rgba(240, 248, 255, 0.8)',
        width=1700,
        height=800
    )

    buffer = io.BytesIO()

    # Export avec une meilleure résolution
    export_fig.write_image(
        buffer,
        format="pdf",
        engine="kaleido",
        scale=2
    )

    pdf_data = buffer.getvalue()
    buffer.close()
    return pdf_data
def display_flight_types(data):
    depart_sec = data[data['Flight_Type'] == 'Depart_Sec']
    night_stop = data[data['Flight_Type'] == 'Night_Stop']

    # Formater l'affichage des dates
    def format_datetime(df):
        df = df.copy()
        df['HA'] = df['HA'].dt.strftime('%H:%M:%S')
        df['HD'] = df['HD'].dt.strftime('%H:%M:%S')
        return df

    col1, col2 = st.columns(2)
    with col1:
        st.write(f"**Nombre de Départs Secs :** {len(depart_sec)}")
        st.dataframe(format_datetime(depart_sec[['VOLD', 'HA', 'HD', 'DEST']]))
    with col2:
        st.write(f"**Nombre de Night Stop :** {len(night_stop)}")
        st.dataframe(format_datetime(night_stop[['VOLA', 'HA', 'HD', 'ORG']]))

st.title("Analyse du programme des vols")

data_file = st.file_uploader("Charger un fichier Excel", type=['xlsx'])

if data_file:
    uploaded_data = pd.ExcelFile(data_file)
    sheet = uploaded_data.sheet_names[0]
    raw_data = uploaded_data.parse(sheet)
    raw_data = cached_preprocess_data(raw_data)

    available_days = raw_data['JOUR'].unique()
    #selected_day = st.selectbox("Sélectionner un jour", available_days)
    #filtered_data = raw_data[raw_data['JOUR'] == selected_day]

    available_dates = pd.to_datetime(raw_data['DATE'].unique())
    selected_date = st.selectbox("Sélectionner une date", available_dates, format_func=lambda x: x.strftime('%d/%m/%Y'))
    filtered_data = raw_data[pd.to_datetime(raw_data['DATE']) == selected_date]

    #st.write("**Données traitées pour le jour sélectionné:**", selected_day)
    if st.checkbox("Afficher/Masquer la liste des vols:"):
        st.dataframe(filtered_data)

    st.write("### Planche des vols")
    gantt_chart = create_interactive_gantt(filtered_data, selected_date)
    st.plotly_chart(gantt_chart)

    st.write("### Types de vols")
    display_flight_types(filtered_data)

    if st.checkbox("Afficher/Masquer les statistiques des vols:"):
        st.write("### Statistiques des vols")
        flights_stats, pax_stats, total_pax, total_flights = calculate_flight_stats(filtered_data)

        col1, col2 = st.columns(2)
        with col1:
            st.write(f"**Nombre de vols par compagnie (Total: {total_flights}):**")
            st.dataframe(flights_stats)
        with col2:
            st.write("**Répartition des vols par compagnie:**")
            components.html(
                create_echarts_html(flights_stats, 'Répartition des vols', 'vols'),
                height=450
            )

        if not pax_stats.empty:
            col3, col4 = st.columns(2)
            with col3:
                st.write(f"**Nombre de passagers par compagnie (Total: {total_pax}):**")
                st.dataframe(pax_stats)
            with col4:
                st.write("**Répartition des passagers par compagnie:**")
                components.html(
                    create_echarts_html(pax_stats, 'Répartition des passagers', 'passagers'),
                    height=450
                )

    export_btn = st.button("Exporter le graphique Gantt en PDF")
    if export_btn:
        try:
            pdf_data = export_gantt_to_pdf(gantt_chart)
            st.download_button(
                label="Télécharger le diagramme Gantt en PDF",
                data=pdf_data,
                file_name=f"planning_vols_{selected_date.strftime('%d_%m_%Y')}.pdf",
                mime="application/pdf"
            )
        except Exception as e:
            st.error(f"Une erreur est survenue lors de l'export du PDF : {str(e)}")
            st.info("Assurez-vous d'avoir installé le package kaleido : pip install kaleido")
