import dash
from dash import dash_table, html, dcc
import pandas as pd
import plotly.graph_objs as go
from dash.dependencies import Output, Input, State
import requests
#import io
import sqlite3
from datetime import datetime, date

#Initialisierung
app = dash.Dash(__name__,
                meta_tags=[{'name': 'viewport', 'content': 'width=device-width, initial-scale=1.0'}])

xls = pd.ExcelFile('wf-dash-copy.xlsx')
df_states = pd.read_excel(xls, sheet_name='states')
df_ressorts = pd.read_excel(xls,sheet_name='ressorts')
df_autor = pd.read_excel(xls,sheet_name='users')
df_data = pd.read_excel(xls, sheet_name='data')

#DB-Connection
def sql_connection():
    try:
        con = sqlite3.connect('workflow-db.sqlite')
        # Tabelle erstellen (falls nicht vorhanden)
        query = """CREATE TABLE if not exists beitraege
        ( id INTEGER PRIMARY KEY AUTOINCREMENT,
        Timeline_Status TEXT, Autor Text, Beitragsthema Text, Ressort TEXT,
        VÖ_Datum DATE, Workflow_Start DATE, Workflow_Ende DATE )"""
        con.execute(query)
        
        empty_check = pd.read_sql_query('SELECT * FROM beitraege', con)

        if empty_check.empty:
            # WICHTIG: Explizit das Blatt 'data' laden!
            df_excel_data = pd.read_excel('wf-dash-copy.xlsx', sheet_name='data')
            # In SQL schreiben
            df_excel_data.to_sql('beitraege', con, index=False, if_exists='append')
            print("Datenbank erfolgreich aus Excel-Blatt 'data' initialisiert.")
        
        con.commit()
        con.close()
    except Exception as e:
        print(f"Fehler beim Laden der DB: {e}")
sql_connection()

def get_df():
    con = sqlite3.connect('workflow-db.sqlite')
    df = pd.read_sql_query('SELECT * FROM beitraege', con)
    con.close()
    
    # Datumsumwandlung für SQL-Standardformat
    for col in ['VÖ_Datum', 'Workflow_Start', 'Workflow_Ende']:
        # KEIN dayfirst=True, da die DB ISO-Format nutzt
        df[col] = pd.to_datetime(df[col], errors='coerce') 
    return df
    
def write_db(data):
    try:
        con = sqlite3.connect('workflow-db.sqlite')
        df_to_save = pd.DataFrame(data)
        for col in ['VÖ_Datum', 'Workflow_Start', 'Workflow_Ende']:
            if col in df_to_save.columns:
                # Hier MUSS dayfirst=True sein, da die Tabelle DD.MM.YYYY sendet
                df_to_save[col] = pd.to_datetime(df_to_save[col], dayfirst=True, errors='coerce')
        # Speichern überschreibt die alte Tabelle mit den sauberen Daten
        df_to_save.to_sql('beitraege', con, if_exists='replace', index=False)
        con.commit()
        con.close()
        return True
    except Exception as e:
        print(f"Fehler: {e}")
        return False

#HTML Components
app.layout = html.Div([
        html.Div([
            html.Img(id='logo', style={'width': '120px'}, src='assets/logo.png'),
            html.H3('MedienMittweida - Beiträge im aktuellen Semester'),
            
            dash_table.DataTable(
                id='table-dropdown',
                data = [],
                columns=[
                     {'id': 'Beitragsthema', 'name': 'Beitragsthema'},
                     {'id': 'Ressort', 'name': 'Ressort','presentation':'dropdown'},
                     {'id': 'Autor', 'name': 'Autor','presentation':'dropdown'},
                     {'id': 'Timeline_Status', 'name': 'Status','presentation':'dropdown'},
                     {'id': 'Workflow_Start', 'name': 'Start'},
                     {'id': 'Workflow_Ende', 'name': 'Ende'},
                     {'id': 'VÖ_Datum', 'name': 'VÖ_Datum'}
                 ],
               
                #(Status-Farben)
                style_data_conditional=[
                    {'if': {'filter_query': '{Timeline_Status} contains "Sheet"'}, 'backgroundColor': "#DDF864", 'color': "#0E141A"},
                    {'if': {'filter_query': '{Timeline_Status} contains "Canva"'}, 'backgroundColor': "#2275D3", 'color': "#D4D4D4"},
                    {'if': {'filter_query': '{Timeline_Status} contains "WP"'}, 'backgroundColor': "#EBB747", 'color': "#313030"},
                    {'if': {'filter_query': '{Timeline_Status} contains "Veröffentlichung"'}, 'backgroundColor': "#13CE39", 'color': "#181A18", 'fontWeight': 'bold'}
                ],
                page_size=8,
                editable=True,
                persistence=True,
                dropdown={
                    'Timeline_Status':{
                    'options':[{'label':i,'value':i} for i in df_states['Timeline_Status'].unique()]},
                    'Ressort':{
                    'options':[{'label':i,'value':i} for i in df_ressorts['Ressort'].unique()]},
                    'Autor':{
                    'options':[{'label':i,'value':i} for i in df_autor['Mitglieder'].unique()]}}
            )
        ], style={'width': '95%', 'margin': '0 auto', 'fontSize': 12}),
        html.Div([
            html.Button('Speichern',id='btn-saved',title='Strg + S funktioniert nicht',
                        n_clicks=0,
                        style={'background':'red','color':'white','padding':'.5rem'}),
            html.Button('Ansicht aktualisieren', id='btn-refresh', n_clicks=0,
                        style={'background':'blue','color':'white','padding':'.5rem'}),
            html.Label(id='state-label')
        ]),

        html.Hr(style={'backgroundColor': 'darkgrey', 'height': '.25rem', 'margin': '20px 0'}),

        html.Div([
            html.Label('Datum wählen für einen bestimmten Zeitraum'),
            dcc.DatePickerRange(
                id='dp',
                display_format='DD.MM.YYYY',
                start_date=date(2025,10,1),
                end_date=date(2027, 12, 31)
            )
        ], style={'textAlign': 'center'}),

        # Erster Graph: Balkendiagramm
        dcc.Graph(id='graph'),
        
        html.Hr(style={'backgroundColor': 'darkgrey', 'height': '.25rem'}),
        
        # Zweiter Graph: Heatmap
        dcc.Graph(id='graph2'),

    ], className='wrapAround', style={'width': '100%', 'margin': '0 auto'})

@app.callback(
    [Output('table-dropdown','data'),
     Output('graph', 'figure'),
     Output('graph2', 'figure'),
     Output('state-label','children')],
    [Input('dp', 'start_date'),
     Input('dp', 'end_date'),
     Input('btn-refresh','n_clicks')]
)
def update_graphs(start, end, n_clicks):
    dff = get_df()
    if dff.empty:
        return [], go.Figure(), go.Figure(), 'Keine Daten'

    dff_filt = dff.sort_values(by='VÖ_Datum', na_position='first')

    # Tabelle für Anzeige formatieren
    table_df = dff_filt.copy()
    for col in ['Workflow_Start', 'Workflow_Ende', 'VÖ_Datum']:
        # Wir formatieren NUR, wenn das Datum NICHT leer ist. 
        # Leere Zellen bleiben einfach leer statt "NaT" anzuzeigen.
        table_df[col] = table_df[col].apply(lambda x: x.strftime('%d.%m.%Y') if pd.notnull(x) else "")

    heute = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    custom_data = dff_filt[['Autor', 'Ressort', 'VÖ_Datum', 'Timeline_Status']]

    # Graph 1 Scatter
    fig1 = go.Figure(go.Bar(
        x=dff_filt['Beitragsthema'],
        y=dff_filt['VÖ_Datum'],
        customdata=custom_data,
        hovertemplate="Status: <br>%{customdata[3]}<br>VÖ-Termin: <br>%{y|%d.%m.%Y}<br><extra></extra>"
    ))
    fig1.update_layout(
        xaxis={'tickangle': -15, 'rangeslider': {'visible': True, 'thickness': 0.15}},
        yaxis={'type': 'date', 'tickformat': '%d.%m.%Y'},
        height=600
    )

    # Graph 2 Heatmap
    z_vals = (dff_filt['VÖ_Datum'] - heute).dt.days
    fig2 = go.Figure(go.Heatmap(
        x=dff_filt['VÖ_Datum'],
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
    
    return table_df.to_dict('records'),fig1, fig2,f"Update: {datetime.now().strftime('%H:%M:%S')}"
# Callback zum Speichern
@app.callback(
    Output('state-label', 'children', allow_duplicate=True),
    Input('btn-saved', 'n_clicks'),
    State('table-dropdown', 'data'),
    prevent_initial_call=True
)
def save_to_db(n_clicks, rows):
    if n_clicks > 0:
        if write_db(rows):
            return "✅ Gespeichert!"
    return dash.no_update

if __name__ == '__main__':
    app.run(debug=True)