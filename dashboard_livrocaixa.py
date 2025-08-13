import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import os
import pandas as pd
import dash
from dash import dcc, html, Input, Output
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta

# ====== CONFIGURAÇÃO ======
arquivo = 'Livro-Caixa.xlsx'
aba = 'Livro Caixa'
META_MENSAL = 100000

# Lê os dados
df = pd.read_excel(arquivo, sheet_name=aba)
df['Data do Pedido'] = pd.to_datetime(df['Data do Pedido'], errors='coerce')
df['AnoMes'] = df['Data do Pedido'].dt.strftime('%Y-%m')
df['DataStr'] = df['Data do Pedido'].dt.strftime('%d/%m')
df['AnoMes'] = df['AnoMes'].astype(str)

meses_validos = [mes for mes in df['AnoMes'].unique() if mes and mes != 'nan']
dropdown_options = [{'label': mes, 'value': mes} for mes in sorted(meses_validos)]

def dias_uteis_restantes(mes_ano):
    hoje = datetime.now().date()
    ano, mes = map(int, mes_ano.split('-'))
    primeiro = datetime(ano, mes, 1).date()
    ultimo = datetime(ano + (mes // 12), (mes % 12) + 1, 1).date() - timedelta(days=1)
    inicio = max(hoje, primeiro) if mes == hoje.month and ano == hoje.year else primeiro
    return int(pd.bdate_range(inicio, ultimo).size)

def generate_client_colors(n):
    import matplotlib.pyplot as plt
    cmap = plt.get_cmap('tab20')
    cores = [f"rgba({int(r*255)},{int(g*255)},{int(b*255)},0.85)" for r,g,b,_ in cmap.colors[:n]]
    return (cores * ((n // len(cores)) + 1))[:n]

app = dash.Dash(__name__)
server = app.server
app.title = "Dashboard Livro Caixa"

app.layout = html.Div(
    style={
        'background': 'radial-gradient(ellipse at center, #181a1e 70%, #121217 100%)',
        'color': '#fafafa',
        'fontFamily': 'Segoe UI, Arial, sans-serif',
        'minHeight': '100vh',
        'minWidth': '100vw',
        'margin': '0',
        'padding': '0',
        'display': 'flex',
        'justifyContent': 'center',
    },
    children=[
        html.Div(
            style={'maxWidth': '1150px', 'width': '95%', 'padding': '20px'},
            children=[
                html.H1("📊 Dashboard Financeiro - Livro Caixa",
                        style={'color': '#00FFB4', 'textAlign': 'center',
                               'marginBottom': '25px', 'textShadow': '0 2px 10px #000'}),
                html.Div([
                    html.Label("Selecione o Mês:",
                               style={'fontWeight': 'bold', 'fontSize': '1.1em',
                                      'marginRight': '10px', 'alignSelf': 'center'}),
                    dcc.Dropdown(id='mes-dropdown', options=dropdown_options,
                                 value=sorted(meses_validos)[-1] if meses_validos else None,
                                 clearable=False,
                                 style={'color': '#000', 'width': '180px'})
                ], style={'display': 'flex', 'justifyContent': 'center',
                          'marginBottom': '30px', 'flexWrap': 'wrap', 'gap': '10px'}),
                html.Div(id='kpis', style={'display': 'flex', 'gap': '18px',
                                           'flexWrap': 'wrap', 'justifyContent': 'center',
                                           'marginBottom': '25px'}),
                html.Div(id='meta-row', style={'marginBottom': '30px',
                                               'marginLeft': 'auto', 'marginRight': 'auto',
                                               'maxWidth': '800px'}),
                dcc.Graph(id='grafico-receita',
                          style={'marginBottom': '25px', 'height': '400px', 'width': '100%'}),
                html.Div([
                    dcc.Graph(id='grafico-comissao',
                              style={'flex': '1 1 400px', 'minWidth': '320px',
                                     'height': '400px', 'maxWidth': '560px'}),
                    dcc.Graph(id='grafico-cliente',
                              style={'flex': '1 1 400px', 'minWidth': '320px',
                                     'height': '400px', 'maxWidth': '560px'})
                ], style={'display': 'flex', 'flexWrap': 'wrap',
                          'justifyContent': 'center', 'gap': '20px'}),
                html.Div("Feito por Feliphe Nunes Silva",
                         style={'color': '#00FFB4', 'textAlign': 'center',
                                'marginTop': '25px', 'fontStyle': 'italic', 'opacity': .7})
            ]
        )
    ]
)

@app.callback(
    Output('kpis', 'children'),
    Output('meta-row', 'children'),
    Output('grafico-receita', 'figure'),
    Output('grafico-comissao', 'figure'),
    Output('grafico-cliente', 'figure'),
    Input('mes-dropdown', 'value')
)
def update_dashboard(mes):
    hoje = datetime.now()
    dados_mes = df[df['AnoMes'] == mes]
    receita = dados_mes['Valor Total'].sum()
    valor_pago = receita
    valor_restante = max(META_MENSAL - valor_pago, 0)
    comissao = dados_mes['Comissão'].sum()
    quantidade = dados_mes['Quantidade'].sum()
    pedidos = len(dados_mes)
    dias_faltam = ((dados_mes['Data do Pedido'].max() or hoje) - hoje).days
    dias_faltam = max(dias_faltam, 0)
    dias_uteis = dias_uteis_restantes(mes)
    meta_diaria_uteis = valor_restante / dias_uteis if dias_uteis > 0 else 0
    meta_percent = (valor_pago / META_MENSAL * 100) if META_MENSAL else 0

    kpi_style = {'background': '#1e2026', 'padding': '16px 20px',
                 'borderRadius': '12px', 'fontSize': '1.05em',
                 'minWidth': '150px', 'textAlign': 'center',
                 'flex': '1 1 120px', 'maxWidth': '240px',
                 'border': '1px solid #333'}
    kpis = [
        html.Div([html.Div("🎯 Meta Mensal", style={'color': '#00FFB4'}),
                  html.Div(f"R$ {META_MENSAL:,.2f}")], style=kpi_style),
        html.Div([html.Div("💸 Valor Pago", style={'color': '#43FF86'}),
                  html.Div(f"R$ {valor_pago:,.2f}")],
                 style={**kpi_style, 'background': 'rgba(67,255,134,0.1)'}),
        html.Div([html.Div("⏳ Valor Restante", style={'color': '#FF1744'}),
                  html.Div(f"R$ {valor_restante:,.2f}")],
                 style={**kpi_style, 'background': 'rgba(255,23,68,0.1)'}),
        html.Div([html.Div("📈 Meta Diária (dias úteis)", style={'color': '#FF7F50'}),
                  html.Div(f"R$ {meta_diaria_uteis:,.2f}/dia"),
                  html.Div(f"({dias_uteis} dias úteis restantes)", style={'fontSize': '0.85em'})],
                 style={**kpi_style, 'background': 'rgba(255,127,80,0.1)'}),
        html.Div([html.Div("📅 Dias p/ final do mês", style={'color': '#7EC8E3'}),
                  html.Div(f"{dias_faltam} dia(s)")],
                 style={**kpi_style, 'background': 'rgba(126,200,227,0.1)'})
    ]

    barra_container = {'position': 'relative', 'width': '100%', 'height': '23px',
                       'borderRadius': '12px', 'background': '#232526',
                       'boxShadow': '0 2px 5px #0008'}
    barra_filled = {'height': '100%', 'width': f"{min(meta_percent, 100):.2f}%",
                    'background': 'linear-gradient(90deg,#00FFB4,#00DD8D 80%)',
                    'borderRadius': '12px 0 0 12px',
                    'transition': 'width .8s cubic-bezier(.25,.1,.4,1.7)'}
    barra_label = {'position': 'absolute', 'top': '0', 'width': '100%',
                   'textAlign': 'center', 'fontWeight': 'bold', 'lineHeight': '23px',
                   'color': '#111'}
    barra_progresso = html.Div(style=barra_container, children=[
        html.Div(style=barra_filled),
        html.Div(f"Progresso da Meta: {meta_percent:.2f}%", style=barra_label)
    ])

    # Gráficos
    fig_receita = px.bar(dados_mes, x='DataStr', y='Valor Total',
                         title='📅 Receita Diária no Mês',
                         labels={'Valor Total': 'Receita (R$)', 'DataStr': 'Data'},
                         color_discrete_sequence=['#00FFB4'],
                         template='plotly_dark')
    fig_receita.update_layout(autosize=False, height=400, width=1150)

    fig_comissao = px.line(dados_mes, x='DataStr', y='Comissão',
                           title='Comissão Diária no Mês',
                           markers=True, template='plotly_dark',
                           color_discrete_sequence=['#FFD740'])
    fig_comissao.update_layout(autosize=False, height=400, width=560)

    if 'Cliente' in dados_mes.columns:
        ranking = dados_mes.groupby('Cliente')['Valor Total'].sum().nlargest(10).reset_index()
    else:
        ranking = pd.DataFrame({'Cliente': ['N/D'], 'Valor Total': [0]})
    cores = generate_client_colors(len(ranking))
    fig_cliente = go.Figure()
    for i, row in ranking.iterrows():
        fig_cliente.add_trace(go.Bar(x=[row['Valor Total']], y=[row['Cliente']],
                                     orientation='h',
                                     marker=dict(color=cores[i], line=dict(color='#232526', width=2)),
                                     hovertemplate='%{y}: R$ %{x:,.2f}<extra></extra>'))
    fig_cliente.update_layout(autosize=False, height=400, width=560,
                              title="🏅 Ranking TOP 10 Clientes",
                              template='plotly_dark',
                              plot_bgcolor='#171C24',
                              paper_bgcolor='rgba(0,0,0,0)')

    return kpis, [barra_progresso], fig_receita, fig_comissao, fig_cliente

# ===== Rodando local ou no Render =====
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 8050))
    app.run(host='0.0.0.0', port=port, debug=True)
