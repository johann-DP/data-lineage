# -*- coding: utf-8 -*-
"""
Created on Tue Jul 11 17:37:35 2023

@author: johan
"""

import networkx as nx
import plotly.graph_objects as go
import dash
from dash import dcc
from dash import html
from dash.dependencies import Input, Output
from bokeh.palettes import Spectral4, Plasma4
from colour import Color

# import du graphe
G = nx.read_gexf("lineage_graph.gexf")

# Création de la liste des fichiers
file_list = list(set([node.split('\n')[0] for node in G.nodes]))
file_list.append('All')

# def convert_color(rgb_color):
#     """Converts an RGB color tuple to a hex string"""
#     return '#%02x%02x%02x' % tuple(int(255 * c) for c in rgb_color)

app = dash.Dash(__name__)
all_columns = list(G.nodes())

# Define some modern colors
background_color = Color("aliceblue").get_hex_l()  
graph_background_color = Color("floralwhite").get_hex_l()  
edge_color_default = Spectral4[0]
edge_color_selected = Spectral4[1]
node_color_default = Plasma4[2]
node_color_selected = Spectral4[3]
node_color_kpi = "#000000" # sns.color_palette("Set2")[0]  


# Layout
app.layout = html.Div([
    html.H1("Data Lineage: Data Visualization", style={"textAlign": "center", "color": "#000000", "font-size": "2.0em", "font-family": "Arial"}),  # Title
    html.Label("Key Performance Indicator", style={"font-size": "1.5em", "font-family": "Calibri", 'margin-left': '15px'}),  # Dropdown title
    dcc.Dropdown(
        id='dropdown',
        options=[{'label': col, 'value': col} for col in all_columns],
        value=all_columns[0],
        style={"width": "50%"}
    ),
    html.Div('File to Display', style={"font-size": "1.5em", "font-family": "Calibri", 'margin-top': '10px', 'margin-bottom': '2px', 'margin-left': '15px'}),
    dcc.Dropdown(
        id='file-dropdown',
        options=[{'label': file, 'value': file} for file in file_list],
        value='All',
        style={"width": "50%"}
    ),
    html.Br(),  # This adds some space
    html.Label("Labels: File - Tabs - KPI", style={"font-size": "1.5em", "font-family": "Calibri", 'margin-left': '15px'}),  # RadioItems title
    dcc.RadioItems(
        id='label-toggle',
        options=[{'label': i, 'value': i} for i in ['Hover', 'Always']],
        value='Hover'
    ),
    html.Br(),  # This adds some space
    dcc.Graph(id='graph'),
    html.Div([
        html.Div('●', style={'color': "#000000", 'display': 'inline-block', 'margin-right': '10px'}),
        html.Span("Selected KPI", style={'display': 'inline-block', 'margin-right': '20px'}),
        html.Div('●', style={'color': node_color_selected, 'display': 'inline-block', 'margin-right': '10px'}),
        html.Span("Selected Node", style={'display': 'inline-block', 'margin-right': '20px'}),
        html.Div('●', style={'color': node_color_default, 'display': 'inline-block', 'margin-right': '10px'}),
        html.Span("Other Nodes", style={'display': 'inline-block', 'margin-right': '20px'}),
        html.Div('—', style={"font-size": "2.5em", 'color': edge_color_default, 'display': 'inline-block', 'margin-right': '10px'}),
        html.Span("KPIs Link", style={'display': 'inline-block', 'margin-right': '20px'}),
        html.Div('—', style={"font-size": "2.5em", 'color': edge_color_selected, 'display': 'inline-block', 'margin-right': '10px'}),
        html.Span("File Relationship", style={'display': 'inline-block'}),
    ], style={'margin-top': '-25px'}),
], style={'backgroundColor': background_color})


@app.callback(
    Output('graph', 'figure'),
    [Input('dropdown', 'value'), 
     Input('graph', 'clickData'),
     Input('label-toggle', 'value'),
     Input('file-dropdown', 'value')]
)
def update_graph(selected_column, click_data, label_toggle, selected_file):
    # Tous les nœuds doivent être affichés par défaut
    nodes_to_display = [node for node in G.nodes() if selected_file == 'All' or node.split('\n')[0] == selected_file]

    if click_data is None:
        selected_node = all_columns[0]
    else:
        if 'text' in click_data['points'][0]:
            selected_node = click_data['points'][0]['text']
        else:
            selected_node = all_columns[0]

    if not G or selected_column not in G.nodes:
        return go.Figure()

    pos = nx.spring_layout(G, seed=42, k=0.3)
    for node in G.nodes:
        G.nodes[node]['pos'] = pos[node]

    edge_traces = []
    for edge in G.edges:
        if edge[0] in nodes_to_display and edge[1] in nodes_to_display:
            x0, y0 = G.nodes[edge[0]]['pos']
            x1, y1 = G.nodes[edge[1]]['pos']
            file_name_source = edge[0].split('\n')[0]
            file_name_target = edge[1].split('\n')[0]
            edge_color = edge_color_selected if file_name_source != file_name_target else edge_color_default
            edge_width = 2.0 if edge[0] == selected_node or edge[1] == selected_node else 0.5
            edge_trace = go.Scatter(
                x=[x0, x1, None], 
                y=[y0, y1, None],
                line=dict(width=edge_width, color=edge_color),
                hoverinfo='none',
                mode='lines'
            )
            edge_traces.append(edge_trace)

    node_x = []
    node_y = []
    node_text = []
    for node in G.nodes():
        x, y = G.nodes[node]['pos']
        node_x.append(x)
        node_y.append(y)
        node_text.append(str(node))

    node_trace = go.Scatter(
        x=node_x, y=node_y,
        text=node_text,
        hoverinfo='text' if label_toggle == 'Hover' else None,
        mode='markers+text' if label_toggle == 'Always' else 'markers',
        marker=dict(
            size=10,
            color=[node_color_selected if node == selected_node else (node_color_kpi if node == selected_column else node_color_default) for node in G.nodes()]
        ),
        textposition='top center'
    )

    fig = go.Figure(data=[*edge_traces, node_trace],
                layout=go.Layout(
                    title='Fruchterman-Reingold Algorithm',
                    showlegend=False,
                    hovermode='closest',
                    margin=dict(b=20, l=5, r=5, t=40),
                    plot_bgcolor=graph_background_color,
                    paper_bgcolor=background_color,
                    height=1000,
                    uirevision='no-change',
                    xaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
                    yaxis=dict(showgrid=False, zeroline=False, showticklabels=False))
                )
    return fig


## DataViz Louis

# from pyvis.network import Network

# def create_html_report2(G, filename, pos=None):
#     # Créer la figure Plotly
#     nt = Network(select_menu=True, filter_menu=True)
#     nt.from_nx(G)
#     nt.show_buttons(filter_=['physics'])
#     nt.save_graph(filename)

# html_file = os.path.join(data_dir, "data_lineage_report.html")
# create_html_report2(G, html_file)

##

if __name__ == '__main__':
    try:
        app.run_server(debug=True) # http://localhost:8050
    except Exception as e:
        print(f"An error occurred: {e}")