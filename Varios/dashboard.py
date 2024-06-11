import dash
from dash import html, dcc
import plotly.graph_objects as go

# Inicializar la aplicación Dash
app = dash.Dash(__name__)

# Datos del ejemplo
resultados = procesar_datos('/path/to/LD PROYECTOS TEMA GENERAL.xlsx')

# Crear un gráfico para cada tipo de evento
app.layout = html.Div([
    html.H1('Dashboard de Eventos por Proyecto'),
    html.Div([
        dcc.Graph(
            id=evento,
            figure={
                'data': [
                    {'x': list(resultados[evento].index), 'y': list(resultados[evento].values), 'type': 'bar', 'name': evento}
                ],
                'layout': {
                    'title': evento
                }
            }
        ) for evento in resultados
    ])
])

# Correr el servidor
if __name__ == '__main__':
    app.run_server(debug=True)
