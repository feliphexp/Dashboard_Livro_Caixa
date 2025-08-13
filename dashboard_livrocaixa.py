import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import pandas as pd
import dash
from dash import dcc, html, Input, Output
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta

# =============== CONFIGURAÇÃO ================
arquivo = 'Livro-Caixa.xlsx'
aba = 'Livro Caixa'
META_MENSAL = 100000

# Leitura e preparação dos dados
df = pd.read_excel(arquivo, sheet_name=aba)
df['Data do Pedido'] = pd.to_datetime(df['Data do Pedido'], errors='coerce')
df['AnoMes'] = df['Data do Pedido'].dt.strftime('%Y-%m')
df['DataStr'] = df['Data do Pedido'].dt.strftime('%d/%m')

# Filtra meses válidos para o dropdown
df['AnoMes'] = df['AnoMes'].astype(str)
meses_validos = [mes for mes in df['AnoMes'].unique() if mes and mes != 'nan']
dropdown_options = [{'label': mes, 'value': mes} for mes in sorted(meses_validos)]

def dias_uteis_restantes(mes_ano):
    hoje = datetime.now().date()
    ano, mes = [int(x) for x in mes_ano.split('-')]
    primeiro = datetime(ano, mes, 1).date()
    if mes == 12:
        ultimo = datetime(ano+1, 1, 1).date() - timedelta(days=1)
    else:
        ultimo = datetime(ano, mes+1, 1).date() - timedelta(days=1)
    inicio = max(hoje, primeiro) if mes == hoje.month and ano == hoje.year else primeiro
    dias_uteis = pd.bdate_range(inicio, ultimo).size
    return int(dias_uteis)

def generate_client_colors(n):
    import matplotlib.pyplot as plt
    cmap = plt.get_cmap('tab20')
    cores_rgba = [f"rgba({int(r*255)},{int(g*255)},{int(b*255)},0.85)" for r,g,b,_ in cmap.colors[:n]]
    if n > len(cmap.colors):
        multiplos = (n // len(cmap.colors)) + 1
        cores_rgba = (cores_rgba * multiplos)[:n]
    return cores_rgba

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
            style={
                'maxWidth': '1150px',
                'width': '95%',
                'margin': '32px 0 40px 0',
                'padding': '0 12px',
                'color': '#fafafa',
                'fontFamily': 'Segoe UI, Arial, sans-serif',
            },
            children=[
                html.H1(
                    "📊 Dashboard Financeiro - Livro Caixa",
                    style={
                        'color': '#00FFB4',
                        'marginBottom': '28px',
                        'textShadow': '0 2px 14px #000',
                        'textAlign': 'center',
                        'fontSize': '2.3em',
                        'fontWeight': 'bold',
                        'userSelect': 'none'
                    }
                ),
                html.Div(
                    style={
                        'display': 'flex',
                        'justifyContent': 'center',
                        'marginBottom': '34px',
                        'flexWrap': 'wrap',
                        'gap': '12px'
                    },
                    children=[
                        html.Label(
                            "Selecione o Mês:",
                            style={
                                'fontWeight': 'bold',
                                'fontSize': '1.22em',
                                'marginRight': '13px',
                                'alignSelf': 'center',
                                'userSelect': 'none'
                            }
                        ),
                        dcc.Dropdown(
                            id='mes-dropdown',
                            options=dropdown_options,
                            value=sorted(meses_validos)[-1] if meses_validos else None,
                            clearable=False,
                            style={
                                'color': '#0a0a0a',
                                'width': '180px',
                                'fontWeight': 'bold',
                            }
                        )
                    ]
                ),
                html.Div(
                    id='kpis',
                    style={
                        'display': 'flex',
                        'gap': '22px',
                        'flexWrap': 'wrap',
                        'justifyContent': 'center',
                        'marginBottom': '30px'
                    }
                ),
                html.Div(
                    id='meta-row',
                    style={
                        'marginBottom': '30px',
                        'marginLeft': 'auto',
                        'marginRight': 'auto',
                        'width': '100%',
                        'maxWidth': '800px'
                    }
                ),
                dcc.Graph(
                    id='grafico-receita',
                    style={'marginBottom': '28px', 'height': '400px', 'width': '100%', 'maxWidth': '1150px'}
                ),
                html.Div(
                    style={
                        'width': '100%',
                        'display': 'flex',
                        'flexWrap': 'wrap',
                        'gap': '24px',
                        'justifyContent': 'center',
                    },
                    children=[
                        dcc.Graph(
                            id='grafico-comissao',
                            style={'flex': '1 1 400px', 'minWidth': '320px', 'height': '400px', 'maxWidth': '570px'}
                        ),
                        dcc.Graph(
                            id='grafico-cliente',
                            style={'flex': '1 1 400px', 'minWidth': '320px', 'height': '400px', 'maxWidth': '570px'}
                        ),
                    ]
                ),
                html.Div(
                    "Feito por Feliphe Nunes Silva",
                    style={
                        'color': '#00FFB4',
                        'textAlign': 'center',
                        'marginTop': '28px',
                        'fontStyle': 'italic',
                        'opacity': 0.7,
                        'fontSize': '0.97em',
                        'userSelect': 'none'
                    }
                )
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
    dados_mes = df[df['AnoMes'] == mes].copy()
    receita = float(dados_mes['Valor Total'].sum())
    valor_pago = receita
    valor_restante = max(META_MENSAL - valor_pago, 0)
    comissao = float(dados_mes['Comissão'].sum())
    quantidade = float(dados_mes['Quantidade'].sum())
    pedidos = len(dados_mes)

    if mes == hoje.strftime('%Y-%m'):
        ultimo_dia = (hoje.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
        dias_faltam = (ultimo_dia - hoje).days
    else:
        data_final = pd.to_datetime(mes + "-01") + pd.offsets.MonthEnd(0)
        dias_faltam = (data_final - pd.to_datetime(mes + "-01")).days

    dias_uteis = dias_uteis_restantes(mes)
    meta_diaria_uteis = valor_restante / dias_uteis if dias_uteis > 0 and valor_restante > 0 else 0
    meta_percent = (valor_pago / META_MENSAL * 100) if META_MENSAL else 0

    kpi_style = {
        'background': 'rgba(25,28,34,0.96)',
        'padding': '18px 22px',
        'borderRadius': '13px',
        'fontSize': '1.12em',
        'minWidth': '150px',
        'fontWeight': 'bold',
        'boxShadow': '0 2px 10px #00FFB422',
        'textAlign': 'center',
        'border': '1.2px solid #223',
        'marginBottom': '7px',
        'flex': '1 1 120px',
        'maxWidth': '240px',
        'userSelect': 'none'
    }

    kpis = [
        html.Div([
            html.Div("🎯 Meta Mensal", style={'color': '#00FFB4', 'fontSize': '1.11em', 'fontWeight': 'bold', 'marginBottom': '6px'}),
            html.Div(f"R$ {META_MENSAL:,.2f}", style={'fontSize': '1.22em'})
        ], style=kpi_style),
        html.Div([
            html.Div("💸 Valor Pago", style={'color': '#43FF86', 'fontWeight': 'bold', 'marginBottom': '6px'}),
            html.Div(f"R$ {valor_pago:,.2f}", style={'fontSize': '1.13em'})
        ], style={**kpi_style,'background': 'rgba(67,255,134,0.08)', 'border': '1.2px solid #43FF86'}),
        html.Div([
            html.Div("⏳ Valor Restante", style={'color': '#FF1744', 'fontWeight': 'bold', 'marginBottom': '6px'}),
            html.Div(f"R$ {valor_restante:,.2f}", style={'fontSize': '1.13em'})
        ], style={**kpi_style, 'background': 'rgba(255,23,68,0.10)', 'border': '1.2px solid #FF1744'}),
        html.Div([
            html.Div("📈 Meta Diária (dias úteis)", style={'color': '#FF7F50', 'fontWeight': 'bold', 'marginBottom': '5px'}),
            html.Div(f"R$ {meta_diaria_uteis:,.2f} / dia", style={'fontSize': '1.07em'}),
            html.Div(f"({dias_uteis} dias úteis restantes)", style={'fontSize': '0.83em', 'color': '#FFA726'})
        ], style={**kpi_style, 'background': 'rgba(255,127,80,0.13)', 'border': '1.2px solid #FF7F50'}),
        html.Div([
            html.Div("⏳ Dias para final do mês", style={'color': '#7EC8E3', 'fontWeight': 'bold', 'marginBottom': '6px'}),
            html.Div(f"{dias_faltam} dia(s)", style={'fontSize': '1.10em', 'fontWeight': 'bold'})
        ], style={**kpi_style, 'background': 'rgba(126,200,227,0.15)', 'border': '1.2px solid #7EC8E3'})
    ]

    barra_container_style = {
        'position': 'relative', 'width': '100%', 'maxWidth': '810px',
        'height': '23px', 'borderRadius': '12px',
        'background': '#232526',
        'margin': 'auto',
        'boxShadow': '0 2px 4px #0008',
        'overflow': 'hidden',
        'userSelect': 'none'
    }
    barra_filled_style = {
        'height': '100%',
        'width': f"{min(meta_percent, 100):.2f}%",
        'background': 'linear-gradient(90deg,#00FFB4,#00DD8D 80%)',
        'borderRadius': '12px 0 0 12px',
        'transition': 'width .8s cubic-bezier(.25,.1,.4,1.7)'
    }
    barra_label_style = {
        'position': 'absolute', 'top': '0', 'left': '0', 'width': '100%', 'height': '100%',
        'textAlign': 'center', 'fontWeight': 'bold', 'lineHeight': '23px',
        'color': '#222', 'textShadow': '0 1.5px 8px #fff6',
        'userSelect': 'none', 'fontSize': '1em'
    }
    barra_progresso = html.Div(style=barra_container_style, children=[
        html.Div(style=barra_filled_style),
        html.Div(f"Progresso da Meta: {meta_percent:.2f}%", style=barra_label_style)
    ])
    meta_row = [barra_progresso]

    fig_receita = px.bar(
        dados_mes.sort_values('Data do Pedido'),
        x='DataStr', y='Valor Total',
        title='📅 Receita Diária no Mês',
        labels={'Valor Total': 'Receita (R$)', 'DataStr': 'Data'},
        color_discrete_sequence=['#00FFB4'],
        template='plotly_dark'
    )
    fig_receita.update_layout(
        autosize=False,
        height=400,
        width=1150,
        font_family='Segoe UI',
        plot_bgcolor='#141516',
        paper_bgcolor='rgba(0,0,0,0)',
        title_font_size=20,
        xaxis_tickangle=-45,
        xaxis_tickvals=dados_mes['DataStr'].unique(),
        xaxis_tickfont=dict(size=10)
    )

    fig_comissao = px.line(
        dados_mes.sort_values('Data do Pedido'),
        x='DataStr', y='Comissão',
        title='Comissão Diária no Mês',
        markers=True,
        labels={'Comissão': 'Comissão (R$)', 'DataStr': 'Data'},
        template='plotly_dark',
        color_discrete_sequence=['#FFD740']
    )
    fig_comissao.update_layout(
        autosize=False,
        height=400,
        width=570,
        font_family='Segoe UI',
        plot_bgcolor='#161818',
        paper_bgcolor='rgba(0,0,0,0)',
        title_font_size=17,
        xaxis_tickangle=-45,
        xaxis_tickvals=dados_mes['DataStr'].unique(),
        xaxis_tickfont=dict(size=10)
    )

    if 'Cliente' in dados_mes.columns:
        ranking = dados_mes.groupby('Cliente')['Valor Total'].sum().sort_values(ascending=False).reset_index()
    else:
        ranking = pd.DataFrame({'Cliente': ['N/D'], 'Valor Total': [0]})

    n_clients = min(len(ranking), 10)
    try:
        cores = generate_client_colors(n_clients)
    except:
        cores = px.colors.qualitative.Bold * (n_clients // len(px.colors.qualitative.Bold) + 1)
        cores = cores[:n_clients]

    fig_cliente = go.Figure()
    for i in range(n_clients):
        fig_cliente.add_trace(go.Bar(
            x=[ranking.loc[i, 'Valor Total']],
            y=[ranking.loc[i, 'Cliente']],
            orientation='h',
            marker=dict(color=cores[i], line=dict(color='#232526', width=2)),
            name=ranking.loc[i, 'Cliente'],
            hovertemplate='%{y}: R$ %{x:,.2f}<extra></extra>'
        ))
    fig_cliente.update_layout(
        autosize=False,
        height=400,
        width=570,
        title="🏅 Ranking TOP 10 Clientes por mês",
        xaxis_title="Valor Comprado (R$)",
        yaxis_title="Cliente",
        font_family='Segoe UI',
        template='plotly_dark',
        plot_bgcolor='#171C24',
        paper_bgcolor='rgba(0,0,0,0)',
        title_font_size=17,
        yaxis={'categoryorder': 'total ascending'},
        barmode='stack',
        showlegend=False,
        margin=dict(l=88, r=15, t=58, b=30)
    )

    return kpis, meta_row, fig_receita, fig_comissao, fig_cliente

if __name__ == '__main__':
    app.run(debug=True)
