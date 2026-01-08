import dash
from dash import dash_table, html, dcc
import pandas as pd
import plotly.graph_objs as go
from dash.dependencies import Output, Input
import requests
import io
from datetime import datetime, date

# 1. Initialisierung
app = dash.Dash(__name__,
                meta_tags=[{'name': 'viewport', 'content': 'width=device-width, initial-scale=1.0'}])

def load_data():
    url_d = 'https://www.dropbox.com/scl/fi/8la5t610wxgklnz1c9n40/wf-dash.xlsx?rlkey=eeu3n3l2jtnp04lyjfeoc3blj&st=jqq7lngu&dl=1'
    try:
        # Kurzer Timeout für Cloud-Stabilität beim Start
        res = requests.get(url_d, timeout=10)
        if res.status_code == 200:
            data_excel = io.BytesIO(res.content)
            df = pd.read_excel(data_excel, sheet_name='data', engine='openpyxl')
            
            # Datum konvertieren
            df['VÖ-Datum'] = pd.to_datetime(df['VÖ-Datum'], errors='coerce')
            # Deine Sortierung: Absteigend (Neueste oben)
            df = df.sort_values(by='VÖ-Datum', ascending=False)
            return df
    except Exception:
        pass
    return pd.DataFrame()

# Layout-Funktion für den automatischen Refresh bei F5
def serve_layout():
    df = load_data()

    # Tabelle-Daten vorbereiten (Datum als Text für deutsche Anzeige)
    table_df = df.copy()
    for col in ['Workflow-Start', 'Workflow-Ende', 'VÖ-Datum']:
        if col in table_df.columns:
            table_df[col] = pd.to_datetime(table_df[col]).dt.strftime('%d.%m.%Y')

    return html.Div([
        html.Div([
            html.Img(id='logo', style={'width': '120px'}, src='assets/logo.png'),
            html.H3('MedienMittweida - Beiträge im aktuellen Semester'),
            dcc.Markdown('''[Link zur Cloud](https://www.dropbox.com/scl/fi/8la5t610wxgklnz1c9n40/wf-dash.xlsx?rlkey=eeu3n3l2jtnp04lyjfeoc3blj&st=jqq7lngu)''',style={'font-size':'1rem'},link_target='_blank'),
            
            dash_table.DataTable(
                id='table-dropdown',
                data=table_df.to_dict('records'),
                columns=[
                    {'id': 'Beitragsthema', 'name': 'Beitragsthema'},
                    {'id': 'Ressort', 'name': 'Ressort'},
                    {'id': 'Autor', 'name': 'Autor'},
                    {'id': 'Timeline-Status', 'name': 'Status'},
                    {'id': 'Workflow-Start', 'name': 'Start'},
                    {'id': 'Workflow-Ende', 'name': 'Ende'},
                    {'id': 'VÖ-Datum', 'name': 'VÖ-Datum'}
                ],
                # DEINE FARB-LOGIK (Status-Farben)
                style_data_conditional=[
                    {'if': {'filter_query': '{Timeline-Status} contains "Sheet"'}, 'backgroundColor': "#DDF864", 'color': "#0E141A"},
                    {'if': {'filter_query': '{Timeline-Status} contains "Canva"'}, 'backgroundColor': "#2275D3", 'color': "#D4D4D4"},
                    {'if': {'filter_query': '{Timeline-Status} contains "WP"'}, 'backgroundColor': "#EBB747", 'color': "#313030"},
                    {'if': {'filter_query': '{Timeline-Status} contains "Veröffentlichung"'}, 'backgroundColor': "#13CE39", 'color': "#181A18", 'fontWeight': 'bold'}
                ],
                page_size=10,
                editable=True
            )
        ], style={'width': '95%', 'margin': '0 auto', 'fontSize': 12}),

        html.Hr(style={'backgroundColor': 'darkgrey', 'height': '.25rem', 'margin': '20px 0'}),

        html.Div([
            html.Label('Datum wählen für einen bestimmten Zeitraum'),
            dcc.DatePickerRange(
                id='dp',
                display_format='DD.MM.YYYY',
                start_date=date(2025,9,1),#df['VÖ-Datum'].min().date(),
                end_date=date(2027, 12, 31)
            )
        ], style={'textAlign': 'center'}),

        # Erster Graph: Balkendiagramm
        dcc.Graph(id='graph'),
        
        html.Hr(style={'backgroundColor': 'darkgrey', 'height': '.25rem'}),
        
        # Zweiter Graph: Heatmap
        dcc.Graph(id='graph2')
        
    ], className='wrapAround', style={'width': '100%', 'margin': '0 auto'})

app.layout = serve_layout

@app.callback(
    [Output('graph', 'figure'),
     Output('graph2', 'figure')],
    [Input('dp', 'start_date'),
     Input('dp', 'end_date')]
)
def update_graphs(start_date, end_date):
    dff = load_data()
    if dff.empty:
        return go.Figure(), go.Figure()
    
    dff['VÖ-Datum-DT'] = pd.to_datetime(dff['VÖ-Datum'])
    mask = (dff['VÖ-Datum-DT'] >= start_date) & (dff['VÖ-Datum-DT'] <= end_date)
    dff_filt = dff.loc[mask].sort_values(by='VÖ-Datum-DT', ascending=True)

    heute = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    custom_data = dff_filt[['Autor', 'Ressort', 'VÖ-Datum', 'Timeline-Status']]

    # Graph 1 (Scatter/Bar)
    fig1 = go.Figure(go.Bar(
        x=dff_filt['Beitragsthema'],
        y=dff_filt['VÖ-Datum-DT'],
        customdata=custom_data,
        hovertemplate="Status: <br>%{customdata[3]}<br>VÖ-Termin: <br>%{y|%d.%m.%Y}<br><extra></extra>"
    ))
    fig1.update_layout(
        xaxis={'tickangle': -15, 'rangeslider': {'visible': True, 'thickness': 0.15}},
        yaxis={'type': 'date', 'tickformat': '%d.%m.%Y'},
        height=600
    )

    # Graph 2 (Heatmap)
    z_vals = (dff_filt['VÖ-Datum-DT'] - heute).dt.days
    fig2 = go.Figure(go.Heatmap(
        x=dff_filt['VÖ-Datum-DT'],
        y=dff_filt['Beitragsthema'],
        z=z_vals,
        colorscale='rdbu',
        showscale=False,
        customdata=custom_data,
        hovertemplate="Thema: %{y}<br>VÖ-Termin: %{x|%d.%m.%Y}<br><extra></extra>"
    ))
    fig2.update_layout(
        xaxis={'type': 'date', 'tickformat': '%d.%m.%Y'},
        margin=dict(l=175, r=25, t=25, b=75),
        height=600,
        shapes=[{'type': 'line', 'x0': heute, 'x1': heute, 'y0': 0, 'y1': 1, 'yref': 'paper', 'line': {'color': 'red', 'width': 5, 'dash': 'dot'}}]
    )
    
    return fig1, fig2

if __name__ == '__main__':
    app.run(debug=False)