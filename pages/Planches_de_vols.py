import pandas as pd
import streamlit as st
from datetime import datetime, timedelta, time
import plotly.express as px
import streamlit.components.v1 as components
import io

st.set_page_config(layout='wide')

@st.cache_data
def cached_preprocess_data(data):
    """Fonction mise en cache pour prétraiter les données."""
    return preprocess_data(data)

def parse_time(time_value):
    if pd.isna(time_value):
        return None
    if isinstance(time_value, str):
        try:
            return datetime.strptime(time_value, '%H:%M:%S')
        except ValueError:
            try:
                return datetime.strptime(time_value, '%H:%M')
            except ValueError:
                return None
    elif isinstance(time_value, datetime):
        return time_value
    elif isinstance(time_value, timedelta):
        return datetime.combine(datetime.today(), (datetime.min + time_value).time())
    elif isinstance(time_value, time):
        return datetime.combine(datetime.today(), time_value)
    return None

def format_date(date_str):
    date_obj = pd.to_datetime(date_str)
    return date_obj.strftime('%d/%m/%Y')

def preprocess_data(data):
    df = data.copy()
    df['HA'] = df['HA'].apply(parse_time)
    df['HD'] = df['HD'].apply(parse_time)
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

    fig = px.timeline(
        data,
        x_start="HA",
        x_end="HD",
        y="Vacation Line",
        width=1500,
        height=600,
        color="Company",
        title=f"Planning des vols du {selected_date.strftime('%d/%m/%Y')}",
        labels={"Company": "Compagnie", "Vacation Line": "Ligne de vacation"},
        hover_data={"HA_str": True, "HD_str": True, "VOLD": True, "VOLA": True, "DEST": True, "ORG": True, "Flight_Type": True}
    )

    for i, row in data.iterrows():
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

        if row['Flight_Type'] == 'Depart_Sec':
            fig.add_shape(
                type="line",
                x0=row['HD'],
                x1=row['HD'],
                y0=row['Vacation Line'] - 0.4,
                y1=row['Vacation Line'] + 0.4,
                line=dict(color="red", width=4),
                layer="above"
            )
        elif row['Flight_Type'] == 'Night_Stop':
            fig.add_shape(
                type="line",
                x0=row['HA'],
                x1=row['HA'],
                y0=row['Vacation Line'] - 0.4,
                y1=row['Vacation Line'] + 0.4,
                line=dict(color="yellow", width=4),
                layer="above"
            )

    fig.update_yaxes(title_text="Lignes de vacation", categoryorder="total ascending")
    fig.update_layout(
        plot_bgcolor='lightblue',
        paper_bgcolor='lightblue',
        margin=dict(l=50, r=50, t=50, b=50),
        showlegend=True
    )

    fig.add_annotation(
        x=1.02,
        y=1,
        xref="paper",
        yref="paper",
        text="Types de vols:<br>Liseret Rouge = Départ Sec<br>Liseret Jaune = Night Stop",
        showarrow=False,
        font=dict(size=12),
        align="left"
    )

    return fig

def export_gantt_to_pdf(fig):
    fig.update_layout(
        paper_bgcolor='rgba(255,255,255,1)',
        plot_bgcolor='rgba(255,255,255,1)',
    )
    buffer = io.BytesIO()
    fig.write_image(
        buffer,
        format="pdf",
        engine="kaleido",
        width=1700,
        height=800,
        scale=2
    )
    pdf_data = buffer.getvalue()
    buffer.close()
    return pdf_data

st.title("Analyse du programme des vols")

data_file = st.file_uploader("Charger un fichier Excel", type=['xlsx'])

if data_file:
    uploaded_data = pd.ExcelFile(data_file)
    sheet = uploaded_data.sheet_names[0]
    raw_data = uploaded_data.parse(sheet)
    raw_data = cached_preprocess_data(raw_data)

    available_days = raw_data['JOUR'].unique()
    selected_day = st.selectbox("Sélectionner un jour", available_days)
    filtered_data = raw_data[raw_data['JOUR'] == selected_day]

    available_dates = pd.to_datetime(raw_data['DATE'].unique())
    selected_date = st.selectbox("Sélectionner une date", available_dates, format_func=lambda x: x.strftime('%d/%m/%Y'))
    filtered_data = raw_data[pd.to_datetime(raw_data['DATE']) == selected_date]

    st.write("**Données traitées pour le jour sélectionné:**", selected_day)
    if st.checkbox("Afficher/Masquer la liste des vols:"):
        st.dataframe(filtered_data)

    st.write("### Planche des vols")
    gantt_chart = create_interactive_gantt(filtered_data, selected_date)
    st.plotly_chart(gantt_chart)

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
