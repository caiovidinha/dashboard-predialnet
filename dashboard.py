import calendar
import datetime
import pandas as pd
import streamlit as st
import plotly.express as px
import re
from pathlib import Path

# --- 1) Carregando os dados do Excel ---
@st.cache_data
def load_data():
    dfs = []
    for f in Path("planilhas").glob("*.xlsx"):
        df = pd.read_excel(f, engine='openpyxl')
        df['Data da Finalização'] = pd.to_datetime(
            df['Data da Finalização'],
            dayfirst=True,
            errors='coerce'
        )
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

data = load_data()

# --- 2) Sidebar: filtros ---
st.sidebar.header("Filtros")

# 2.1) Filtro por mês
meses_map = {
    1:  'Janeiro',
    2:  'Fevereiro',
    3:  'Março',
    4:  'Abril',
    5:  'Maio',
    6:  'Junho',
    7:  'Julho',
    8:  'Agosto',
    9:  'Setembro',
    10: 'Outubro',
    11: 'Novembro',
    12: 'Dezembro',
}

# 2.1) Filtro por meses com multiselect (dropdown)
available_months = sorted(data['Data da Finalização']
                          .dt.month
                          .dropna()
                          .unique()
                          .astype(int))
month_names = [meses_map[m] for m in available_months]

sel_meses = []
with st.sidebar.expander("Meses", expanded=False):
    for nome in month_names:
        if st.checkbox(nome, value=True, key=f"mes_{nome}"):
            sel_meses.append(nome)

# 2.2) Filtros extras de data
excluir_sabados = st.sidebar.checkbox("Excluir sábados", value=False)
limitar_dias_totais = st.sidebar.checkbox(
    "Dias Totais (ajusta todos os meses ao menor dia comum)",
    value=False
)

# 2.3) Filtros de Plano/Cidade/Bairro
planos  = ['Todos'] + sorted(data['Plano'].dropna().unique().tolist())
cidades = ['Todos'] + sorted(data['Cidade'].dropna().unique().tolist())
bairros = ['Todos'] + sorted(data['Bairro'].dropna().unique().tolist())

f_plano  = st.sidebar.selectbox("Plano",  planos,  index=0)
f_cidade = st.sidebar.selectbox("Cidade", cidades, index=0)
f_bairro = st.sidebar.selectbox("Bairro", bairros, index=0)


# --- 3) Aplica filtro de mês sobre o dataset inicial ---
df = data.copy()
if sel_meses:
    # converte nomes de volta para números
    nums = [k for k,v in meses_map.items() if v in sel_meses]
    df = df[df['Data da Finalização'].dt.month.isin(nums)]

# --- 4) Prepara df_dates com filtros de data ---
df_dates = df.copy()

if excluir_sabados:
    df_dates = df_dates[df_dates['Data da Finalização'].dt.weekday != 5]

if limitar_dias_totais:
    last_days = (
        df
        .groupby(df['Data da Finalização'].dt.to_period('M'))['Data da Finalização']
        .max()
        .dt.day
    )
    cutoff_day = int(last_days.min())
    df_dates = df_dates[df_dates['Data da Finalização'].dt.day <= cutoff_day]

# conta única de dias válidos
dias_validos = df_dates['Data da Finalização'].dt.date.nunique()

# --- 5) Aplica filtros de Plano/Cidade/Bairro ---
df = df_dates.copy()
if f_plano  != 'Todos':
    df = df[df['Plano']  == f_plano]
if f_cidade != 'Todos':
    df = df[df['Cidade'] == f_cidade]
if f_bairro != 'Todos':
    df = df[df['Bairro'] == f_bairro]
# depois de aplicar todos os filtros em `df`
max_bairros = df['Bairro'].nunique()
top_n = st.sidebar.slider(
    "Número de bairros a exibir",
    min_value=1,
    max_value=max_bairros,
    value=min(10, max_bairros),
    step=1
)

# --- 5) Layout e Métricas Gerais ---
st.title("📊 Dashboard de Agendamentos")
st.markdown("### 📈 Métricas Gerais")

col1, col2, col3 = st.columns(3)
col1.metric("Total de Registros", f"{len(df)}")
col2.metric("Dias no Período", f"{dias_validos}")
media_diaria = str(round((len(df) / dias_validos),2))
col3.metric("Média diária", f"{media_diaria.replace('.',',')}")

# --- 6) Gráficos de Distribuição ---
st.markdown("### Gráficos de Distribuição")

# Distribuição por Plano
df_plano = df.groupby('Plano').size().reset_index(name='Quantidade')
fig_plano = px.bar(
    df_plano, x='Plano', y='Quantidade',
    text='Quantidade', title="Distribuição por Plano",
    height=600
)
fig_plano.update_traces(texttemplate='%{text}', textposition='outside')
st.plotly_chart(fig_plano, use_container_width=True)

# --- Gráfico Foco: Planos 500, 800 e 1Gb – todos juntos numa só barra cada ---
st.markdown("### Distribuição Principal: Planos 500 / 800 / 1Gb")

def categoriza_foco(plano: str) -> str | None:
    p = plano.lower()
    # 500 fica 500
    if re.search(r'\b500\b', p):
        return "1- 500 Mega"
    # 600 ou 800 ficam 800
    if re.search(r'\b600\b', p) or re.search(r'\b800\b', p):
        return "2- 800 Mega (600 mega até abril)"
    # 700 ou 1gb ficam 1 Giga
    if re.search(r'\b700\b', p) or re.search(r'\b1\s?gb\b', p):
        return "3- 1 Giga (700 mega até abril)"
    return None


# Adiciona coluna de planos principais
df['PlanosPrincipais'] = df['Plano'].apply(categoriza_foco)
df_foco = df[df['PlanosPrincipais'].notnull()]
df_foco_count = (
    df_foco
    .groupby('PlanosPrincipais')
    .size()
    .reset_index(name='Quantidade')
    .sort_values('Quantidade', ascending=False)
)

# Calcula porcentagem de cada plano principal
total_foco = df_foco_count['Quantidade'].sum()
df_foco_count['%'] = (df_foco_count['Quantidade'] / total_foco * 100).round(1)
df_foco_count['label'] = df_foco_count.apply(lambda r: f"{r['PlanosPrincipais']} ({r['%']}%)", axis=1)

fig_foco = px.bar(
    df_foco_count,
    x='label', y='Quantidade',
    text='Quantidade',
    title="Agendamentos nos Planos Principais",
    height=600
)
fig_foco.update_traces(texttemplate='%{text}', textposition='outside')
fig_foco.update_layout(xaxis_title="Plano (% do total)")
st.plotly_chart(fig_foco, use_container_width=True)


# Distribuição por Cidade
df_cidade = df.groupby('Cidade').size().reset_index(name='Quantidade')
fig_cidade = px.bar(
    df_cidade, x='Cidade', y='Quantidade',
    text='Quantidade', title="Distribuição por Cidade",
    height=600
)
fig_cidade.update_traces(texttemplate='%{text}', textposition='outside')
st.plotly_chart(fig_cidade, use_container_width=True)

# Distribuição por Bairro (vertical, ordenado, barras maiores e scroll)
df_bairro = (
    df
    .groupby('Bairro')
    .size()
    .reset_index(name='Quantidade')
    .sort_values('Quantidade', ascending=False)
).head(top_n)

bar_width = 80
chart_width = len(df_bairro) * bar_width

fig_bairro = px.bar(
    df_bairro,
    x='Bairro', y='Quantidade',
    text='Quantidade',
    title="Distribuição por Bairro",
    height=600,
    width=chart_width        # força scroll horizontal
)
fig_bairro.update_traces(texttemplate='%{text}', textposition='outside')
fig_bairro.update_layout(
    xaxis={'categoryorder':'total descending'},
    margin={'b':200},         # espaço pra labels inclinados
    xaxis_tickangle=-45
)

st.plotly_chart(fig_bairro, use_container_width=False)


df_tech = (
    df
    .groupby('Tecnologia')
    .size()
    .reset_index(name='Quantidade')
    .sort_values('Quantidade', ascending=False)
)

fig_tech = px.bar(
    df_tech,
    x='Tecnologia', y='Quantidade',
    text='Quantidade',
    title="Distribuição por Tecnologia",
    height=600
)
fig_tech.update_traces(texttemplate='%{text}', textposition='outside')
# Se tiver muitos valores, você pode rotacionar os rótulos:
fig_tech.update_layout(xaxis_tickangle=-45, margin={'b':200})

st.plotly_chart(fig_tech, use_container_width=True)

# --- 7) Análise temporal ---
st.markdown("### Análise por Dia do Mês")
df['Dia'] = df['Data da Finalização'].dt.day
agr = df.groupby('Dia').size().reset_index(name='Contagem')
fig_tempo = px.line(
    agr, x='Dia', y='Contagem',
    markers=True, title="Agendamentos por Dia do Mês"
)
st.plotly_chart(fig_tempo, use_container_width=True)

# Média de agendamentos por Dia do Mês
# df_date = (
#     df.groupby(df['Data da Finalização'].dt.date)
#       .size()
#       .reset_index(name='Contagem')
# )
# df_date['Dia'] = pd.to_datetime(df_date['Data da Finalização']).dt.day
# media_por_dia = (
#     df_date
#     .groupby('Dia')['Contagem']
#     .mean()
#     .round(2)
#     .reset_index(name='Média')
# )
# st.markdown("### Média de agendamentos por Dia do Mês")
# st.dataframe(media_por_dia.set_index('Dia'), use_container_width=True)

# --- 8) Tabela completa com filtros ---
df_disp = (
    df[['Data da Finalização','Bairro','Cidade','Plano','Tecnologia']]
      # converte a data para string DD/MM/AA
      .assign(**{
          'Data da Finalização': lambda d: d['Data da Finalização']
                                          .dt.strftime('%d/%m/%y')
      })
      .sort_values('Data da Finalização', ascending=False)
      .set_index('Data da Finalização')
)

st.markdown("### Tabela de Registros")
st.dataframe(df_disp, use_container_width=True)

# --- Bloco Extra: Identificação de Meses Incompletos e Previsão ---
# st.markdown("## Projeção de Agendamentos por Mês (feriados ignorados para passado)")

df_pred = df_dates.copy()
df_pred['Year']  = df_pred['Data da Finalização'].dt.year
df_pred['Month'] = df_pred['Data da Finalização'].dt.month

agg = (
    df_pred
    .groupby(['Year','Month'])
    .agg(
        registros=('Data da Finalização','count'),
        dias_coletados=('Data da Finalização', lambda x: x.dt.date.nunique())
    )
    .reset_index()
)

# calcula dias sem domingo
def count_non_sundays(year, month):
    start = pd.Timestamp(year=year, month=month, day=1)
    end   = start + pd.offsets.MonthEnd(0)
    days  = pd.date_range(start, end, freq='D')
    return int((days.weekday != 6).sum())

agg['dias_no_mes'] = agg.apply(lambda r: count_non_sundays(r.Year, r.Month), axis=1)

# identifica mês/ano atual
today = datetime.date.today()
current_year  = today.year
current_month = today.month

# força meses passados a completos (ignora feriados)
mask_passado = (
    (agg['Year'] < current_year) |
    ((agg['Year'] == current_year) & (agg['Month']  < current_month))
)
# para esses, dias_coletados = dias_no_mes
agg.loc[mask_passado, 'dias_coletados'] = agg.loc[mask_passado, 'dias_no_mes']

# status
agg['Status'] = agg.apply(
    lambda r: '✅' if r.dias_coletados >= r.dias_no_mes else '⌛',
    axis=1
)

# média e previsão
agg['media_diaria']   = (agg['registros'] / agg['dias_coletados']).round(2)
agg['previsao_total'] = (
    agg
    .apply(lambda r: r.registros if mask_passado.loc[r.name] else round(r.media_diaria * r.dias_no_mes), 
           axis=1)
).astype(int)

# formata média e mês/ano
agg['media_diaria'] = agg['media_diaria'].map(lambda x: f"{x:.2f}".replace('.', ','))
agg['Mês/Ano'] = agg.apply(lambda r: f"{meses_map[r.Month]}/{int(r.Year)}", axis=1)

# monta tabela
tabela_pred = agg[[
    'Mês/Ano','Status','registros','dias_coletados','dias_no_mes',
    'media_diaria','previsao_total'
]].rename(columns={
    'registros': 'Registros até hoje',
    'dias_coletados': 'Dias Coletados',
    'dias_no_mes': 'Dias Válidos',
    'media_diaria': 'Média Diária',
    'previsao_total': 'Previsão Total'
})

# st.dataframe(tabela_pred.set_index('Mês/Ano'), use_container_width=True)

# --- 10) Download dos dados filtrados ---
csv = df.to_csv(index=False).encode('utf-8')
st.download_button(
    "📥 Baixar dados filtrados",
    data=csv,
    file_name='agendamentos_filtrados.csv',
    mime='text/csv'
)
