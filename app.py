import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
import plotly.graph_objects as go
import pywhatkit
import time
from datetime import datetime
import pyautogui
import os

# Configuração da página Streamlit
st.set_page_config(page_title="Dashboard de Faturamento", layout="wide")

# Aplicando o tema escuro e melhorias visuais
st.markdown("""
<style>
    /* Tema escuro geral */
    body {
        color: #E0E0E0;
        background-color: #121212;
    }

    /* Estilo para os títulos */
    h1, h2, h3 {
        color: #BB86FC;
    }

    /* Estilo para os cards de métricas */
    .stMetric {
        background-color: #1E1E1E;
        border: 1px solid #333333;
        border-radius: 5px;
        padding: 10px;
    }

    /* Estilo para os expanders */
    .streamlit-expanderHeader {
        background-color: #1E1E1E;
        color: #BB86FC;
    }

    /* Estilo para as tabelas */
    .dataframe {
        background-color: #1E1E1E;
        color: #E0E0E0;
    }

    /* Estilo para os sliders e selectboxes */
    .stSlider, .stSelectbox {
        background-color: #1E1E1E;
        color: #E0E0E0;
    }

    /* Estilo para os botões */
    .stButton > button {
        background-color: #BB86FC;
        color: #121212;
    }

    /* Estilo para as barras de progresso */
    .stProgress > div > div {
        background-color: #03DAC6;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center;'>Dashboard | TR. Petry</h1>", unsafe_allow_html=True)

# Função para carregar os dados do Excel
@st.cache_data
def load_data():
    return pd.read_excel(r"C:\Users\Petry\Documents\Py\Relatorio MDFe (51).xlsx")

# Carregamento e pré-processamento dos dados
Tabela_MDFE_original = load_data()
Tabela_MDFE_original['Data de emissão'] = pd.to_datetime(Tabela_MDFE_original['Data de emissão'], format='%d/%m/%Y')

# Opção para mostrar/ocultar filtros
show_filters = st.checkbox("Mostrar filtros", value=False)

# Lógica de filtros
if show_filters:
    # Adicionando filtros
    col_placa, col_situacao = st.columns(2)

    with col_placa:
        placas = ["Todos"] + sorted(Tabela_MDFE_original['Placa(s)'].unique().tolist())
        placa_filtro = st.multiselect("Filtrar por placa", placas, default=["Todos"])

    with col_situacao:
        situacoes = sorted(Tabela_MDFE_original['Situação'].unique().tolist())
        situacao_filtro = st.multiselect("Filtrar por situação", situacoes, default=["Encerrada", "Autorizada"])
else:
    # Valores padrão quando os filtros estão ocultos
    placa_filtro = ["Todos"]
    situacao_filtro = ["Encerrada", "Autorizada"]

# Aplicação dos filtros
Tabela_MDFE = Tabela_MDFE_original.copy()
if "Todos" not in placa_filtro:
    Tabela_MDFE = Tabela_MDFE[Tabela_MDFE['Placa(s)'].isin(placa_filtro)]
Tabela_MDFE = Tabela_MDFE[Tabela_MDFE['Situação'].isin(situacao_filtro)]

# Definições de datas e períodos
data_atual = datetime.now()
mes_atual = data_atual.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
data_12_meses_atras = mes_atual - pd.DateOffset(months=12)
ultimo_dia_faturamento = Tabela_MDFE[Tabela_MDFE['Data de emissão'] >= mes_atual]['Data de emissão'].max()
mes_anterior = mes_atual - pd.DateOffset(months=1)
inicio_mes_anterior = mes_anterior
fim_mes_anterior = ultimo_dia_faturamento - pd.DateOffset(months=1)

# Cálculos de faturamento
dados_mes_atual = Tabela_MDFE[(Tabela_MDFE['Data de emissão'] >= mes_atual) & (Tabela_MDFE['Data de emissão'] <= ultimo_dia_faturamento)]
dados_mes_anterior = Tabela_MDFE[(Tabela_MDFE['Data de emissão'] >= inicio_mes_anterior) & (Tabela_MDFE['Data de emissão'] <= fim_mes_anterior)]
faturamento_mes_atual = dados_mes_atual['Valor do(s) CTe(s) vinculado(s)'].sum()

# Cálculo do faturamento dos últimos 12 meses para a linha de tendência
faturamento_12_meses_tendencia = []
for i in range(12):
    data_inicio = mes_atual - pd.DateOffset(months=i+1)
    data_fim = data_inicio + pd.DateOffset(months=1) - pd.Timedelta(days=1)
    data_fim = min(data_fim, data_atual)
    faturamento = Tabela_MDFE[
        (Tabela_MDFE['Data de emissão'] >= data_inicio) & 
        (Tabela_MDFE['Data de emissão'] <= data_fim)
    ]['Valor do(s) CTe(s) vinculado(s)'].sum()
    faturamento_12_meses_tendencia.append(faturamento)

media_12_meses_tendencia = sum(faturamento_12_meses_tendencia) / 12

faturamento_12_meses = Tabela_MDFE[Tabela_MDFE['Data de emissão'] >= data_12_meses_atras]['Valor do(s) CTe(s) vinculado(s)'].sum()

# Função para calcular o imposto do Simples Nacional
def calcular_imposto_simples(faturamento_12_meses, faturamento_mes_atual):
    # Anexo III da Lei Complementar 123/2006 (atualizada)
    faixas = [
        (0, 180000, 0.06, 0),
        (180000.01, 360000, 0.1112, 9360),
        (360000.01, 720000, 0.1340, 17640),
        (720000.01, 1800000, 0.1608, 35640),
        (1800000.01, 3600000, 0.1912, 71280),
        (3600000.01, 4800000, 0.2100, 143280)
    ]

    for limite_inferior, limite_superior, aliquota, deducao in faixas:
        if faturamento_12_meses <= limite_superior:
            break
    
    if faturamento_12_meses > 4800000:
        raise ValueError("Faturamento excede o limite do Simples Nacional")

    aliquota_efetiva = ((faturamento_12_meses * aliquota) - deducao) / faturamento_12_meses
    imposto = faturamento_mes_atual * aliquota_efetiva

    return aliquota_efetiva, imposto

# Função para calcular métricas de faturamento
def calcular_metricas(dados_atual, dados_anterior, situacao_filtro, meta_mensal=100000.00, meta_diaria=5000.00, meta_peso=600.0, meta_semanal=25000.00):
    # Aplicar o filtro de situação
    dados_atual_filtrado = dados_atual[dados_atual['Situação'].isin(situacao_filtro)]
    dados_anterior_filtrado = dados_anterior[dados_anterior['Situação'].isin(situacao_filtro)]
    
    faturamento_rota_atual = dados_atual_filtrado.groupby('Rota')['Valor do(s) CTe(s) vinculado(s)'].sum()
    faturamento_rota_anterior = dados_anterior_filtrado.groupby('Rota')['Valor do(s) CTe(s) vinculado(s)'].sum()
    
    faturamento = dados_atual_filtrado['Valor do(s) CTe(s) vinculado(s)'].sum()
    dias_uteis = dados_atual_filtrado[dados_atual_filtrado['Data de emissão'].dt.dayofweek < 5]
    media_diaria = dias_uteis.groupby('Data de emissão')['Valor do(s) CTe(s) vinculado(s)'].sum().mean()
    peso_total = dados_atual_filtrado['Peso Bruto'].sum() / 1000
    ultimo_dia = dados_atual_filtrado['Data de emissão'].dt.date.max()
    faturamento_ultimo_dia = dados_atual_filtrado[dados_atual_filtrado['Data de emissão'].dt.date == ultimo_dia]['Valor do(s) CTe(s) vinculado(s)'].sum()
    data_atual = pd.Timestamp.now().date()
    inicio_semana = data_atual - pd.Timedelta(days=data_atual.weekday())
    dados_semana = dados_atual_filtrado[(dados_atual_filtrado['Data de emissão'].dt.date >= inicio_semana) & (dados_atual_filtrado['Situação'].isin(['Encerrada', 'Autorizada']))]
    faturamento_semana = dados_semana['Valor do(s) CTe(s) vinculado(s)'].sum()
    
    faturamento_placa_atual = dados_atual_filtrado.groupby('Placa(s)')['Valor do(s) CTe(s) vinculado(s)'].sum()
    faturamento_placa_anterior = dados_anterior_filtrado.groupby('Placa(s)')['Valor do(s) CTe(s) vinculado(s)'].sum()
    
    return {
        'faturamento': faturamento,
        'media_diaria': media_diaria,
        'peso_total': peso_total,
        'faturamento_ultimo_dia': faturamento_ultimo_dia,
        'faturamento_semana': faturamento_semana,
        'faturamento_rota_atual': faturamento_rota_atual,
        'faturamento_rota_anterior': faturamento_rota_anterior,
        'faturamento_placa_atual': faturamento_placa_atual,
        'faturamento_placa_anterior': faturamento_placa_anterior,
        'meta_mensal': meta_mensal,
        'meta_diaria': meta_diaria,
        'meta_peso': meta_peso,
        'meta_semanal': meta_semanal,
        'ultimo_dia': ultimo_dia,
        'inicio_semana': inicio_semana,
        'data_atual': data_atual,
        'media_12_meses_tendencia': media_12_meses_tendencia
    }

# Cálculo do imposto e métricas
aliquota_efetiva, imposto = calcular_imposto_simples(faturamento_12_meses, faturamento_mes_atual)
metricas = calcular_metricas(dados_mes_atual, dados_mes_anterior, situacao_filtro)

# Função para criar uma barra de progresso personalizada com cor padrão
def custom_progress_bar(percentage, text):
    original_font_size = 16
    new_font_size = original_font_size * 0.85

    # Definindo uma cor padrão para todas as barras
    color = "#03DAC6"  # Cor de destaque do tema escuro

    progress_html = f"""
    <div style="
        width: 100%;
        height: 20px;
        background-color: #1E1E1E;
        border-radius: 10px;
        overflow: hidden;
    ">
        <div style="
            width: {min(percentage, 100)}%;
            height: 100%;
            background-color: {color};
            display: flex;
            align-items: center;
            justify-content: center;
        ">
            <span style="
                color: #121212;
                font-weight: bold;
                font-size: {new_font_size}px;
            ">{text}</span>
        </div>
    </div>
    """
    st.markdown(progress_html, unsafe_allow_html=True)

# Função para formatar valores monetários no padrão brasileiro
def formatar_moeda_br(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# Layout principal do dashboard
col1, col2, col3, col4, col5, col6 = st.columns(6)

# Componentes do dashboard (métricas, gráficos, etc.)
with col3:
    st.metric("Faturamento Mensal", formatar_moeda_br(metricas['faturamento']))
    porcentagem_meta_mensal = metricas['faturamento'] / metricas['meta_mensal'] * 100
    custom_progress_bar(porcentagem_meta_mensal, f"{porcentagem_meta_mensal:.1f}%")
    
    # Adiciona espaço de 5px
    st.markdown('<div style="margin-bottom: 5px;"></div>', unsafe_allow_html=True)
    
    # Tabela detalhes por placa
    with st.expander("Detalhes por Placa"):
        # Ordenar as placas por faturamento atual
        placas_ordenadas = sorted(metricas['faturamento_placa_atual'].items(), key=lambda x: x[1], reverse=True)

        # Definir a meta semanal por placa
        meta_semanal_placa = 20000.00

        # Criando cards para cada placa
        for placa, faturamento_atual in placas_ordenadas:
            # Criando um container para o card
            card_container = st.container()
            with card_container:
                # Título: 3 primeiros caracteres da placa + faturamento
                titulo_card = f"{placa[:3]} - R$ {faturamento_atual:,.2f}"
                st.markdown(f"<h4 style='margin-bottom: 5px; font-size: 1em;'>{titulo_card}</h4>", unsafe_allow_html=True)
                
                porcentagem_meta = min(faturamento_atual / meta_semanal_placa * 100, 100)
                custom_progress_bar(porcentagem_meta, f"{porcentagem_meta:.1f}%")
            
            # Adicionando uma barra de divisão entre os cards
            st.markdown('<hr style="border:none; height:1px; background-color:#e0e0e0; margin:5px 0;">', unsafe_allow_html=True)

    # Após o expander "Detalhes por Placa", adicione este novo expander:

    # Tabela detalhes por rota
    with st.expander("Detalhes por Rota"):
        # Calcular o faturamento por rota
        faturamento_rota = dados_mes_atual.groupby('Rota')['Valor do(s) CTe(s) vinculado(s)'].sum()
        
        # Ordenar as rotas por faturamento
        rotas_ordenadas = sorted(faturamento_rota.items(), key=lambda x: x[1], reverse=True)

        # Definir metas específicas para cada rota
        metas_rota = {
            "CD CAXIAS": 50000,
            "GRAMADO": 20000,
            "PORTO ALEGRE": 20000,
            "FAZENDA": 10000
        }

        # Criando cards para cada rota
        for rota, faturamento_atual in rotas_ordenadas:
            # Criando um container para o card
            card_container = st.container()
            with card_container:
                # Título: nome da rota + faturamento
                titulo_card = f"{rota} - R$ {faturamento_atual:,.2f}"
                st.markdown(f"<h4 style='margin-bottom: 5px; font-size: 1em;'>{titulo_card}</h4>", unsafe_allow_html=True)
                
                # Obter a meta específica para a rota ou usar um valor padrão se não estiver definida
                meta_rota = metas_rota.get(rota, 20000)  # Valor padrão de 20000 se a rota não estiver no dicionário
                
                porcentagem_meta = min(faturamento_atual / meta_rota * 100, 100)
                custom_progress_bar(porcentagem_meta, f"{porcentagem_meta:.1f}%")
            
            # Adicionando uma barra de divisão entre os cards
            st.markdown('<hr style="border:none; height:1px; background-color:#e0e0e0; margin:5px 0;">', unsafe_allow_html=True)

with col4:
    st.metric("Peso Total", f"{metricas['peso_total']:.2f} t")
    porcentagem_meta_peso = metricas['peso_total'] / metricas['meta_peso'] * 100
    custom_progress_bar(porcentagem_meta_peso, f"{porcentagem_meta_peso:.1f}%")

    # Adiciona espaço de 5px
    st.markdown('<div style="margin-bottom: 5px;"></div>', unsafe_allow_html=True)

    # Detalhes de Peso por Placa
    with st.expander("Detalhes de Peso por Placa"):
        # Calcular o peso por placa
        peso_placa = dados_mes_atual.groupby('Placa(s)')['Peso Bruto'].sum() / 1000  # Convertendo para toneladas
        
        # Ordenar as placas por peso
        placas_ordenadas = sorted(peso_placa.items(), key=lambda x: x[1], reverse=True)

        # Definir a meta de peso por placa
        meta_peso_placa = 120.0  # toneladas

        # Criando cards para cada placa
        for placa, peso_atual in placas_ordenadas:
            card_container = st.container()
            with card_container:
                titulo_card = f"{placa[:3]} - {peso_atual:.2f} t"
                st.markdown(f"<h4 style='margin-bottom: 5px; font-size: 1em;'>{titulo_card}</h4>", unsafe_allow_html=True)
                
                porcentagem_meta = min(peso_atual / meta_peso_placa * 100, 100)
                custom_progress_bar(porcentagem_meta, f"{porcentagem_meta:.1f}%")
            
            st.markdown('<hr style="border:none; height:1px; background-color:#e0e0e0; margin:5px 0;">', unsafe_allow_html=True)

    # Detalhes de Peso por Rota
    with st.expander("Detalhes de Peso por Rota"):
        # Calcular o peso por rota
        peso_rota = dados_mes_atual.groupby('Rota')['Peso Bruto'].sum() / 1000  # Convertendo para toneladas
        
        # Ordenar as rotas por peso
        rotas_ordenadas = sorted(peso_rota.items(), key=lambda x: x[1], reverse=True)

        # Definir metas específicas de peso para cada rota
        metas_peso_rota = {
            "CD CAXIAS": 250.0,
            "GRAMADO": 140.0,
            "PORTO ALEGRE": 140.0,
            "FAZENDA": 70.0
        }

        # Criando cards para cada rota
        for rota, peso_atual in rotas_ordenadas:
            card_container = st.container()
            with card_container:
                titulo_card = f"{rota} - {peso_atual:.2f} t"
                st.markdown(f"<h4 style='margin-bottom: 5px; font-size: 1em;'>{titulo_card}</h4>", unsafe_allow_html=True)
                
                meta_rota = metas_peso_rota.get(rota, 140.0)  # Valor padrão de 140 toneladas se a rota não estiver no dicionário
                porcentagem_meta = min(peso_atual / meta_rota * 100, 100)
                custom_progress_bar(porcentagem_meta, f"{porcentagem_meta:.1f}%")
            
            st.markdown('<hr style="border:none; height:1px; background-color:#e0e0e0; margin:5px 0;">', unsafe_allow_html=True)

with col5:
    st.metric("Média Diária", formatar_moeda_br(metricas['media_diaria']))
    porcentagem_meta_diaria = metricas['media_diaria'] / metricas['meta_diaria'] * 100
    custom_progress_bar(porcentagem_meta_diaria, f"{porcentagem_meta_diaria:.1f}%")

    # Adiciona espaço de 5px
    st.markdown('<div style="margin-bottom: 5px;"></div>', unsafe_allow_html=True)

    # Detalhes da Média Diária
    with st.expander("Detalhes da Média Diária"):
        # Calcular as médias diárias dos últimos 5 dias
        ultimos_5_dias = dados_mes_atual['Data de emissão'].dt.date.unique()[-5:]
        medias_diarias = []

        faturamento_acumulado = 0
        dias_acumulados = 0

        for dia in dados_mes_atual['Data de emissão'].dt.date.unique():
            dados_dia = dados_mes_atual[dados_mes_atual['Data de emissão'].dt.date == dia]
            faturamento_dia = dados_dia['Valor do(s) CTe(s) vinculado(s)'].sum()
            faturamento_acumulado += faturamento_dia
            dias_acumulados += 1
            media_diaria = faturamento_acumulado / dias_acumulados

            if dia in ultimos_5_dias:
                medias_diarias.append((dia, media_diaria))

        # Criar tabela com as médias diárias e setas indicativas
        df_medias = pd.DataFrame(medias_diarias, columns=['Data', 'Média Diária'])
        df_medias['Data'] = df_medias['Data'].apply(lambda x: x.strftime('%d/%m'))  # Formatação alterada para dia/mês
        
        # Calcular a variação e adicionar setas
        df_medias['Variação'] = df_medias['Média Diária'].pct_change() * 100
        df_medias['Seta'] = df_medias['Variação'].apply(lambda x: "↑" if x >= 0 else "↓")
        df_medias['Cor'] = df_medias['Variação'].apply(lambda x: "green" if x >= 0 else "red")
        
        # Formatar a média diária e a variação
        df_medias['Média Diária'] = df_medias['Média Diária'].apply(lambda x: f'R$ {x:,.2f}')
        df_medias['Variação'] = df_medias['Variação'].apply(lambda x: f'{abs(x):.2f}%' if pd.notnull(x) else '')
        
        # Inverter a ordem para mostrar o dia mais recente primeiro
        df_medias = df_medias.iloc[::-1].reset_index(drop=True)

        # Criar uma tabela HTML personalizada
        html_table = "<table style='width:100%'>"
        html_table += "<tr><th>Data</th><th>Média Diária</th><th>Variação</th></tr>"
        
        for _, row in df_medias.iterrows():
            html_table += f"<tr>"
            html_table += f"<td>{row['Data']}</td>"
            html_table += f"<td>{row['Média Diária']}</td>"
            html_table += f"<td><span style='color:{row['Cor']}'>{row['Seta']} {row['Variação']}</span></td>"
            html_table += f"</tr>"
        
        html_table += "</table>"

        # Exibir a tabela HTML
        st.markdown(html_table, unsafe_allow_html=True)

        # Comparar com a média do dia anterior (mantido para referência)
        if len(medias_diarias) >= 2:
            media_atual = medias_diarias[-1][1]
            media_anterior = medias_diarias[-2][1]
            variacao = (media_atual - media_anterior) / media_anterior * 100
            seta = "↑" if variacao >= 0 else "↓"
            cor = "green" if variacao >= 0 else "red"
            st.markdown(f"<p style='color: {cor};'>{seta} {abs(variacao):.2f}% em relação ao dia anterior</p>", unsafe_allow_html=True)

    # Nova tabela para as últimas 5 semanas
    with st.expander("Detalhes da Média Semanal"):
        # Calcular as médias semanais das últimas 5 semanas
        dados_mes_atual['Semana'] = dados_mes_atual['Data de emissão'].dt.to_period('W')
        ultimas_5_semanas = dados_mes_atual['Semana'].unique()[-5:]
        medias_semanais = []

        for semana in ultimas_5_semanas:
            dados_semana = dados_mes_atual[dados_mes_atual['Semana'] == semana]
            faturamento_semana = dados_semana['Valor do(s) CTe(s) vinculado(s)'].sum()
            dias_uteis = dados_semana[dados_semana['Data de emissão'].dt.dayofweek < 5]['Data de emissão'].nunique()
            media_diaria_semana = faturamento_semana / dias_uteis if dias_uteis > 0 else 0
            medias_semanais.append((semana.start_time, media_diaria_semana))

        # Criar tabela com as médias semanais e setas indicativas
        df_medias_semanais = pd.DataFrame(medias_semanais, columns=['Data', 'Média Diária'])
        df_medias_semanais['Data'] = df_medias_semanais['Data'].apply(lambda x: x.strftime('%d/%m'))
        
        # Calcular a variação e adicionar setas
        df_medias_semanais['Variação'] = df_medias_semanais['Média Diária'].pct_change() * 100
        df_medias_semanais['Seta'] = df_medias_semanais['Variação'].apply(lambda x: "↑" if x >= 0 else "↓")
        df_medias_semanais['Cor'] = df_medias_semanais['Variação'].apply(lambda x: "green" if x >= 0 else "red")
        
        # Formatar a média diária e a variação
        df_medias_semanais['Média Diária'] = df_medias_semanais['Média Diária'].apply(lambda x: f'R$ {x:,.2f}')
        df_medias_semanais['Variação'] = df_medias_semanais['Variação'].apply(lambda x: f'{abs(x):.2f}%' if pd.notnull(x) else '')
        
        # Inverter a ordem para mostrar a semana mais recente primeiro
        df_medias_semanais = df_medias_semanais.iloc[::-1].reset_index(drop=True)

        # Criar uma tabela HTML personalizada
        html_table = "<table style='width:100%'>"
        html_table += "<tr><th>Semana</th><th>Média Diária</th><th>Variação</th></tr>"
        
        for _, row in df_medias_semanais.iterrows():
            html_table += f"<tr>"
            html_table += f"<td>{row['Data']}</td>"
            html_table += f"<td>{row['Média Diária']}</td>"
            html_table += f"<td><span style='color:{row['Cor']}'>{row['Seta']} {row['Variação']}</span></td>"
            html_table += f"</tr>"
        
        html_table += "</table>"

        # Exibir a tabela HTML
        st.markdown(html_table, unsafe_allow_html=True)

        # Comparar com a média da semana anterior
        if len(medias_semanais) >= 2:
            media_atual = medias_semanais[-1][1]
            media_anterior = medias_semanais[-2][1]
            variacao = (media_atual - media_anterior) / media_anterior * 100
            seta = "↑" if variacao >= 0 else "↓"
            cor = "green" if variacao >= 0 else "red"
            st.markdown(f"<p style='color: {cor};'>{seta} {abs(variacao):.2f}% em relação à semana anterior</p>", unsafe_allow_html=True)

# Card do imposto do Simples Nacional
with col6:
    st.metric(f"Imposto Simples Nacional ({aliquota_efetiva:.2%})", formatar_moeda_br(imposto))
    porcentagem_imposto = (imposto / faturamento_mes_atual) * 100
    custom_progress_bar(porcentagem_imposto, f"{porcentagem_imposto:.1f}%")
    
    # Adiciona espaço de 5px
    st.markdown('<div style="margin-bottom: 5px;"></div>', unsafe_allow_html=True)
    
    # Tabela com informações de imposto dos últimos meses
    with st.expander("Detalhes do Imposto"):
        # Calcula o imposto para os últimos 6 meses
        meses_anteriores = [mes_atual - pd.DateOffset(months=i) for i in range(6)]
        impostos_anteriores = []
        for data_inicio in meses_anteriores:
            data_fim = data_inicio + pd.DateOffset(months=1) - pd.Timedelta(days=1)
            faturamento_mes = Tabela_MDFE[(Tabela_MDFE['Data de emissão'] >= data_inicio) & (Tabela_MDFE['Data de emissão'] <= data_fim)]['Valor do(s) CTe(s) vinculado(s)'].sum()
            _, imposto_mes = calcular_imposto_simples(faturamento_12_meses, faturamento_mes)
            impostos_anteriores.append({'Data': data_inicio, 'Imposto': imposto_mes})
        
        df_impostos = pd.DataFrame(impostos_anteriores)
        df_impostos['Data'] = df_impostos['Data'].dt.strftime('%m/%Y')
        df_impostos['Imposto'] = df_impostos['Imposto'].apply(lambda x: f'R$ {x:,.2f}')
        
        st.table(df_impostos)

    # Novo expander para "Detalhes do Imposto 2"
    with st.expander("Detalhes do Imposto 2"):
        # Calcula o imposto para os últimos 12 meses
        meses_12_anteriores = [mes_atual - pd.DateOffset(months=i) for i in range(12)]
        impostos_12_meses = []
        for data_inicio in meses_12_anteriores:
            data_fim = data_inicio + pd.DateOffset(months=1) - pd.Timedelta(days=1)
            faturamento_mes = Tabela_MDFE[(Tabela_MDFE['Data de emissão'] >= data_inicio) & (Tabela_MDFE['Data de emissão'] <= data_fim)]['Valor do(s) CTe(s) vinculado(s)'].sum()
            aliquota_efetiva_mes, imposto_mes = calcular_imposto_simples(faturamento_12_meses, faturamento_mes)
            impostos_12_meses.append({
                'Data': data_inicio,
                'Faturamento': faturamento_mes,
                'Alíquota Efetiva': aliquota_efetiva_mes,
                'Imposto': imposto_mes
            })
        
        df_impostos_12_meses = pd.DataFrame(impostos_12_meses)
        df_impostos_12_meses['Data'] = df_impostos_12_meses['Data'].dt.strftime('%m/%Y')
        df_impostos_12_meses['Faturamento'] = df_impostos_12_meses['Faturamento'].apply(lambda x: f'R$ {x:,.2f}')
        df_impostos_12_meses['Alíquota Efetiva'] = df_impostos_12_meses['Alíquota Efetiva'].apply(lambda x: f'{x:.2%}')
        df_impostos_12_meses['Imposto'] = df_impostos_12_meses['Imposto'].apply(lambda x: f'R$ {x:,.2f}')
        
        st.table(df_impostos_12_meses)

# Card de faturamento do último dia
with col1:
    # Criando um dicionário para mapear os nomes dos meses em português
    meses_pt = {
        1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
        5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
        9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
    }
    
    # Formatando a data por extenso em português
    ultimo_dia_formatado = f"{metricas['ultimo_dia'].day} de {meses_pt[metricas['ultimo_dia'].month]}"
    
    st.metric(f"{ultimo_dia_formatado}", formatar_moeda_br(metricas['faturamento_ultimo_dia']))
    porcentagem_meta_diaria = metricas['faturamento_ultimo_dia'] / metricas['meta_diaria'] * 100
    custom_progress_bar(porcentagem_meta_diaria, f"{porcentagem_meta_diaria:.1f}%")
    
    # Adiciona espaço de 5px
    st.markdown('<div style="margin-bottom: 5px;"></div>', unsafe_allow_html=True)
    
    # Tabela detalhes ultimo dia
    with st.expander("Detalhes do Último Dia"):
        # Filtra os dados do último dia
        ultimo_dia = dados_mes_atual['Data de emissão'].dt.date.max()
        dados_ultimo_dia = dados_mes_atual[dados_mes_atual['Data de emissão'].dt.date == ultimo_dia]

        # Cria um DataFrame com as colunas desejadas
        df_ultimo_dia = dados_ultimo_dia[['Placa(s)', 'Valor do(s) CTe(s) vinculado(s)', 'Rota', 'Peso Bruto']]
        df_ultimo_dia = df_ultimo_dia.rename(columns={
            'Placa(s)': 'Placa',
            'Valor do(s) CTe(s) vinculado(s)': 'Fat',
            'Peso Bruto': 'Peso'
        })

        # Formata a coluna de placa para mostrar apenas os 3 primeiros caracteres
        df_ultimo_dia['Placa'] = df_ultimo_dia['Placa'].str[:3]

        # Substitui "CD CAXIAS" por "CD" na coluna Rota
        df_ultimo_dia['Rota'] = df_ultimo_dia['Rota'].replace("CD CAXIAS", "CD")

        # Ordena o DataFrame pelo faturamento em ordem decrescente
        df_ultimo_dia = df_ultimo_dia.sort_values('Fat', ascending=False)

        # Formata a coluna de faturamento
        df_ultimo_dia['Fat'] = df_ultimo_dia['Fat'].apply(lambda x: f'R$ {x:,.2f}')

        # Formata a coluna de peso (convertendo para toneladas)
        df_ultimo_dia['Peso'] = df_ultimo_dia['Peso'].apply(lambda x: f"{x/1000:.2f} t")

        # Reordena as colunas
        df_ultimo_dia = df_ultimo_dia[['Placa', 'Fat', 'Rota', 'Peso']]

        # Exibe a tabela usando o método padrão do Streamlit, ocultando a coluna de índice
        st.dataframe(df_ultimo_dia, use_container_width=True, hide_index=True)

    # Nova tabela para os últimos 5 dias
    with st.expander("Detalhes dos últimos 5 dias"):
        # Filtra os dados dos últimos 5 dias
        ultimos_5_dias = dados_mes_atual['Data de emissão'].dt.date.unique()[-5:]
        dados_ultimos_5_dias = dados_mes_atual[dados_mes_atual['Data de emissão'].dt.date.isin(ultimos_5_dias)]

        # Cria um DataFrame com as colunas desejadas
        df_ultimos_5_dias = dados_ultimos_5_dias[['Data de emissão', 'Placa(s)', 'Valor do(s) CTe(s) vinculado(s)', 'Rota', 'Peso Bruto']]
        df_ultimos_5_dias = df_ultimos_5_dias.rename(columns={
            'Data de emissão': 'Data',
            'Placa(s)': 'Placa',
            'Valor do(s) CTe(s) vinculado(s)': 'Fat',
            'Peso Bruto': 'Peso'
        })

        # Formata a coluna de data
        df_ultimos_5_dias['Data'] = df_ultimos_5_dias['Data'].dt.strftime('%d/%m')

        # Formata a coluna de placa para mostrar apenas os 3 primeiros caracteres
        df_ultimos_5_dias['Placa'] = df_ultimos_5_dias['Placa'].str[:3]

        # Substitui "CD CAXIAS" por "CD" na coluna Rota
        df_ultimos_5_dias['Rota'] = df_ultimos_5_dias['Rota'].replace("CD CAXIAS", "CD")

        # Ordena o DataFrame pela data (mais recente primeiro) e pelo faturamento em ordem decrescente
        df_ultimos_5_dias = df_ultimos_5_dias.sort_values(['Data', 'Fat'], ascending=[False, False])

        # Formata a coluna de faturamento
        df_ultimos_5_dias['Fat'] = df_ultimos_5_dias['Fat'].apply(lambda x: f'R$ {x:,.2f}')

        # Formata a coluna de peso (convertendo para toneladas)
        df_ultimos_5_dias['Peso'] = df_ultimos_5_dias['Peso'].apply(lambda x: f"{x/1000:.2f} t")

        # Reordena as colunas
        df_ultimos_5_dias = df_ultimos_5_dias[['Data', 'Placa', 'Fat', 'Rota', 'Peso']]

        # Exibe a tabela usando o método padrão do Streamlit, ocultando a coluna de índice
        st.dataframe(df_ultimos_5_dias, use_container_width=True, hide_index=True)

# Card faturamento semana
with col2:
    # Calculando o último dia da semana (considerando que pode ser a data atual)
    fim_semana = min(metricas['inicio_semana'] + pd.Timedelta(days=6), metricas['data_atual'])
    
    st.metric("Faturamento Semanal", formatar_moeda_br(metricas['faturamento_semana']))
   
    porcentagem_meta_semanal = metricas['faturamento_semana'] / metricas['meta_semanal'] * 100
    custom_progress_bar(porcentagem_meta_semanal, f"{porcentagem_meta_semanal:.1f}%")
    
    # Adiciona espaço de 5px
    st.markdown('<div style="margin-bottom: 5px;"></div>', unsafe_allow_html=True)
    
    # Tabela detalhes da semana
    with st.expander("Total semana Peso"):
        # Filtra os dados da semana atual
        dados_semana = dados_mes_atual[
            (dados_mes_atual['Data de emissão'].dt.date >= metricas['inicio_semana']) & 
            (dados_mes_atual['Data de emissão'].dt.date <= fim_semana)
        ]
        
        # Calcula o peso total da semana
        peso_total_semana = dados_semana['Peso Bruto'].sum() / 1000  # Convertendo para toneladas
        
        # Cria um DataFrame com as colunas desejadas
        df_semana = dados_semana.groupby('Placa(s)').agg({
            'Valor do(s) CTe(s) vinculado(s)': 'sum',
            'Peso Bruto': 'sum'
        }).reset_index()
        
        # Renomeia as colunas
        df_semana.columns = ['Placa', 'Fat', 'Peso']
        
        # Formata a coluna de placa para mostrar apenas os 3 primeiros caracteres
        df_semana['Placa'] = df_semana['Placa'].str[:3]
        
        # Ordena o DataFrame pelo faturamento em ordem decrescente
        df_semana = df_semana.sort_values('Fat', ascending=False)
        
        # Formata a coluna de faturamento
        df_semana['Fat'] = df_semana['Fat'].apply(lambda x: f'R$ {x:,.2f}')
        
        # Formata a coluna de peso (convertendo para toneladas)
        df_semana['Peso'] = df_semana['Peso'].apply(lambda x: f"{x/1000:.2f} t")
        
        # Exibe a tabela usando o método padrão do Streamlit, ocultando a coluna de índice
        st.dataframe(df_semana, use_container_width=True, hide_index=True)
    
    # Novo botão de detalhes com tabela por rota
    with st.expander("Total semana Rota"):
        # Cria um DataFrame com as colunas desejadas
        df_semana_rota = dados_semana.groupby(['Placa(s)', 'Rota']).agg({
            'Valor do(s) CTe(s) vinculado(s)': 'sum'
        }).reset_index()
        
        # Renomeia as colunas
        df_semana_rota.columns = ['Placa', 'Rota', 'Fat']
        
        # Formata a coluna de placa para mostrar apenas os 3 primeiros caracteres
        df_semana_rota['Placa'] = df_semana_rota['Placa'].str[:3]
        
        # Substitui "CD CAXIAS" por "CD" na coluna Rota
        df_semana_rota['Rota'] = df_semana_rota['Rota'].replace("CD CAXIAS", "CD")
        
        # Ordena o DataFrame pelo faturamento em ordem decrescente
        df_semana_rota = df_semana_rota.sort_values('Fat', ascending=False)
        
        # Formata a coluna de faturamento
        df_semana_rota['Fat'] = df_semana_rota['Fat'].apply(lambda x: f'R$ {x:,.2f}')
        
        # Exibe a tabela usando o método padrão do Streamlit, ocultando a coluna de índice
        st.dataframe(df_semana_rota, use_container_width=True, hide_index=True)

st.markdown('<hr style="border:none; height:1px; background-color:#e0e0e0; margin:5px 0;">', unsafe_allow_html=True)

# Adicionando um checkbox para controlar a visibilidade da tabela e filtros
show_table_and_filters = st.checkbox("Mostrar tabela e filtros detalhados", value=False)

# Movendo a tabela para uma nova linha
if show_table_and_filters:
    col_tabela = st.columns(1)[0]

    with col_tabela:
        # Criando a tabela com as colunas especificadas
        colunas_tabela = ['Data de emissão', 'Situação', 'Valor do(s) CTe(s) vinculado(s)', 'Peso Bruto', 'Placa(s)', 'Rota']
        df_tabela = Tabela_MDFE_original[colunas_tabela].copy()

        # Convertendo a coluna 'Data de emissão' para datetime
        df_tabela['Data de emissão'] = pd.to_datetime(df_tabela['Data de emissão'])

        # Obtendo a data mínima e máxima do DataFrame
        data_min = df_tabela['Data de emissão'].min().date()
        data_max = df_tabela['Data de emissão'].max().date()

        # Criando um seletor de intervalo de datas usando st.slider
        col1, col2, col3 = st.columns(3)
        with col1:
            intervalo_datas = st.slider(
                "Selecione o intervalo de datas",
                min_value=data_min,
                max_value=data_max,
                value=(data_min, data_max),
                format="DD/MM/YYYY"
            )
            data_inicio, data_fim = intervalo_datas
        with col2:
            situacoes = st.multiselect("Situação", options=df_tabela['Situação'].unique(), default=['Encerrada', 'Autorizada'])
        with col3:
            placas = st.multiselect("Placa(s)", options=['Todas'] + list(df_tabela['Placa(s)'].unique()), default=['Todas'])

        # Filtrando o DataFrame com base nas datas selecionadas, situações e placas
        df_filtrado = df_tabela[
            (df_tabela['Data de emissão'].dt.date >= data_inicio) & 
            (df_tabela['Data de emissão'].dt.date <= data_fim) &
            (df_tabela['Situação'].isin(situacoes))
        ]

        # Aplicando o filtro de placas
        if 'Todas' not in placas:
            df_filtrado = df_filtrado[df_filtrado['Placa(s)'].isin(placas)]

        # Calculando os totais do DataFrame filtrado
        total_valor_cte = df_filtrado['Valor do(s) CTe(s) vinculado(s)'].sum()
        total_peso_bruto = df_filtrado['Peso Bruto'].sum()

        # Formatando os valores monetários e de peso
        df_filtrado['Valor do(s) CTe(s) vinculado(s)'] = df_filtrado['Valor do(s) CTe(s) vinculado(s)'].apply(formatar_moeda_br)
        df_filtrado['Peso Bruto'] = df_filtrado['Peso Bruto'].apply(lambda x: f"{x:,.2f} kg")

        # Criando o cabeçalho com os totais
        st.markdown(f"**Total Valor do(s) CTe(s) vinculado(s): {formatar_moeda_br(total_valor_cte)}**")
        st.markdown(f"**Total Peso Bruto: {total_peso_bruto:,.2f} kg**")

        # Exibindo a tabela com os registros filtrados
        st.dataframe(df_filtrado, use_container_width=True)

# Modificando as colunas para que os gráficos ocupem 50% do espaço cada
col_grafico_placa, col_grafico_rota = st.columns(2)

with col_grafico_placa:
    df_faturamento_placa = pd.DataFrame({
        'Mês Atual': metricas['faturamento_placa_atual'],
        'Mês Anterior': metricas['faturamento_placa_anterior']
    }).reset_index().fillna(0).sort_values('Mês Atual', ascending=True)  # Alterado para ascending=True
    
    fig_faturamento_placa = go.Figure()
    
    # Adicionando barras verticais para o Mês Atual
    fig_faturamento_placa.add_trace(go.Bar(
        x=df_faturamento_placa['Placa(s)'],
        y=df_faturamento_placa['Mês Atual'],
        name='Mês Atual',
        marker_color='#636EFA'
    ))
    
    # Mantendo a linha para o Mês Anterior
    fig_faturamento_placa.add_trace(go.Scatter(
        x=df_faturamento_placa['Placa(s)'], 
        y=df_faturamento_placa['Mês Anterior'], 
        mode='lines+markers',
        name='Mês Anterior', 
        line=dict(color='red', width=3),
        marker=dict(symbol='diamond', size=8, color='red'),
        opacity=0.7
    ))
    
    fig_faturamento_placa.update_layout(
        title='Faturamento por Placa', 
        height=500,
        margin=dict(t=50, b=50, l=30, r=30), 
        font=dict(family="Arial, sans-serif", size=10, color="#E0E0E0"),  # Adicionada família de fonte
        legend=dict(
            orientation="h", 
            yanchor="top", 
            y=1.02, 
            xanchor="right", 
            x=1,
            font=dict(family="Arial, sans-serif", size=12)  # Fonte da legenda
        ),
        xaxis=dict(
            tickangle=45, 
            title='', 
            title_standoff=22, 
            tickmode='array', 
            tickvals=list(range(len(df_faturamento_placa['Placa(s)']))), 
            ticktext=df_faturamento_placa['Placa(s)'], 
            tickfont=dict(family="Arial, sans-serif", size=8),  # Fonte dos ticks do eixo x
            autorange='reversed'
        ),
        barmode='group',
        plot_bgcolor='#1E1E1E',
        paper_bgcolor='#121212',
        title_font=dict(family="Arial, sans-serif", size=14)  # Fonte do título
    )
    st.plotly_chart(fig_faturamento_placa, use_container_width=True)

with col_grafico_rota:
    def formatar_reais(valor):
        return f'R$ {valor:,.2f}'.replace(',', '_').replace('.', ',').replace('_', '.')

    df_faturamento_rota = pd.DataFrame({
        'Mês Atual': metricas['faturamento_rota_atual'],
        'Mês Anterior': metricas['faturamento_rota_anterior']
    }).reset_index().fillna(0).sort_values('Mês Atual', ascending=True)
    
    fig_faturamento_rota = go.Figure()
    fig_faturamento_rota.add_trace(go.Bar(
        x=df_faturamento_rota['Rota'], y=df_faturamento_rota['Mês Atual'], name='Mês Atual',
        text=[formatar_reais(valor) for valor in df_faturamento_rota['Mês Atual']],
        textposition='outside', marker_color='#636EFA', orientation='v', textfont=dict(size=10)
    ))
    fig_faturamento_rota.add_trace(go.Scatter(
        x=df_faturamento_rota['Rota'], y=df_faturamento_rota['Mês Anterior'], mode='lines+markers',
        name='Mês Anterior', line=dict(color='red', width=3), marker=dict(symbol='diamond', size=8, color='red'),
        opacity=0.7
    ))
    fig_faturamento_rota.update_layout(
        title='Faturamento por Rota', 
        height=500,
        margin=dict(t=50, b=50, l=30, r=30), 
        font=dict(family="Arial, sans-serif", size=10, color="#E0E0E0"),  # Adicionada família de fonte
        legend=dict(
            orientation="h", 
            yanchor="top", 
            y=1.02, 
            xanchor="right", 
            x=1,
            font=dict(family="Arial, sans-serif", size=12)  # Fonte da legenda
        ),
        xaxis=dict(
            tickangle=0, 
            title='', 
            title_standoff=22, 
            tickfont=dict(family="Arial, sans-serif", size=10)  # Fonte dos ticks do eixo x
        ),
        plot_bgcolor='#1E1E1E',
        paper_bgcolor='#121212',
        title_font=dict(family="Arial, sans-serif", size=14)  # Fonte do título
    )
    st.plotly_chart(fig_faturamento_rota, use_container_width=True)

print(f"Período atual: {mes_atual} até {data_atual}")
print(f"Período anterior: {inicio_mes_anterior} até {fim_mes_anterior}")
print(f"Registros mês atual: {len(dados_mes_atual)}")
print(f"Registros mês anterior: {len(dados_mes_anterior)}")
print(f"Faturamento mês atual: {dados_mes_atual['Valor do(s) CTe(s) vinculado(s)'].sum()}")
print(f"Faturamento mês anterior: {dados_mes_anterior['Valor do(s) CTe(s) vinculado(s)'].sum()}")

# Função para criar o relatório de texto principal
def criar_relatorio_texto(metricas):
    relatorio = f"""Relatório de Faturamento - {datetime.now().strftime('%d/%m/%Y %H:%M')}

Faturamento Mensal: {formatar_moeda_br(metricas['faturamento'])}
Peso Total: {metricas['peso_total']:.2f} t
Média Diária: {formatar_moeda_br(metricas['media_diaria'])}
Faturamento Semanal: {formatar_moeda_br(metricas['faturamento_semana'])}

Faturamento do último dia ({metricas['ultimo_dia'].strftime('%d/%m')}): {formatar_moeda_br(metricas['faturamento_ultimo_dia'])}

Top 3 Placas por Faturamento:
"""
    top_placas = sorted(metricas['faturamento_placa_atual'].items(), key=lambda x: x[1], reverse=True)[:3]
    for placa, faturamento in top_placas:
        relatorio += f"- {placa[:3]}: {formatar_moeda_br(faturamento)}\n"
    
    relatorio += "\nTop 3 Rotas por Faturamento:\n"
    top_rotas = sorted(metricas['faturamento_rota_atual'].items(), key=lambda x: x[1], reverse=True)[:3]
    for rota, faturamento in top_rotas:
        relatorio += f"- {rota}: {formatar_moeda_br(faturamento)}\n"
    
    return relatorio

# Nova função para criar o relatório detalhado do último dia
def criar_relatorio_ultimo_dia(metricas, dados_mes_atual):
    ultimo_dia = dados_mes_atual['Data de emissão'].dt.date.max()
    dados_ultimo_dia = dados_mes_atual[dados_mes_atual['Data de emissão'].dt.date == ultimo_dia]
    
    relatorio = f"""Detalhes do Último Dia de Faturamento ({ultimo_dia.strftime('%d/%m/%Y')})

Faturamento Total: {formatar_moeda_br(metricas['faturamento_ultimo_dia'])}
Número de Viagens: {len(dados_ultimo_dia)}
Peso Total: {dados_ultimo_dia['Peso Bruto'].sum() / 1000:.2f} t

Detalhes por Placa:
"""
    # Criar um DataFrame com as colunas desejadas
    df_ultimo_dia = dados_ultimo_dia[['Placa(s)', 'Valor do(s) CTe(s) vinculado(s)', 'Rota', 'Peso Bruto']]
    df_ultimo_dia = df_ultimo_dia.rename(columns={
        'Placa(s)': 'Placa',
        'Valor do(s) CTe(s) vinculado(s)': 'Fat',
        'Peso Bruto': 'Peso'
    })

    # Formatar e ordenar os dados
    df_ultimo_dia['Placa'] = df_ultimo_dia['Placa'].str[:3]
    df_ultimo_dia['Rota'] = df_ultimo_dia['Rota'].replace("CD CAXIAS", "CD")
    df_ultimo_dia = df_ultimo_dia.sort_values('Fat', ascending=False)
    
    for _, row in df_ultimo_dia.iterrows():
        relatorio += f"- {row['Placa']}: {formatar_moeda_br(row['Fat'])} | {row['Rota']} | {row['Peso']/1000:.2f} t\n"
    
    return relatorio

# Função para enviar mensagem pelo WhatsApp
def enviar_whatsapp(numero_destino, mensagem):
    try:
        pywhatkit.sendwhatmsg_instantly(
            numero_destino,
            mensagem,
            wait_time=10,
            tab_close=True
        )
        print("Mensagem enviada com sucesso!")
        return True
    except Exception as e:
        print(f"Erro ao enviar mensagem: {str(e)}")
        return False

# Função para capturar a primeira página do dashboard
def capturar_primeira_pagina():
    # Espera um pouco para garantir que a página esteja carregada
    time.sleep(2)
    
    # Obtém as dimensões da tela
    screen_width, screen_height = pyautogui.size()
    
    # Define as coordenadas para capturar apenas o conteúdo da guia
    # Ajuste esses valores conforme necessário para seu monitor
    left = 0  # Borda esquerda da tela
    top = 100  # Altura estimada da barra de título do navegador
    right = screen_width  # Borda direita da tela
    bottom = screen_height - 40  # Altura estimada da barra de tarefas
    
    # Captura a região especificada
    screenshot = pyautogui.screenshot(region=(left, top, right - left, bottom - top))
    
    # Cria um diretório para salvar as capturas de tela, se não existir
    if not os.path.exists("screenshots"):
        os.makedirs("screenshots")
    
    # Gera um nome de arquivo único baseado na data e hora atual
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"screenshots/dashboard_primeira_pagina_{timestamp}.png"
    
    # Salva a captura de tela
    screenshot.save(filename)
    
    return filename

# Função para fechar o navegador
def fechar_navegador():
    # Pressiona Alt+F4 para fechar a janela ativa (navegador)
    pyautogui.hotkey('alt', 'f4')

# Modificar o bloco de envio do relatório
col_envio = st.columns(1)[0]
with col_envio:
    if st.button("Enviar Relatório por WhatsApp"):
        numero_whatsapp = "+55519766-2816"  # Número fixo para envio
        st.info("Preparando relatórios e captura de tela... Por favor, aguarde.")
        
        # Criar relatórios de texto
        relatorio_principal = criar_relatorio_texto(metricas)
        relatorio_ultimo_dia = criar_relatorio_ultimo_dia(metricas, dados_mes_atual)
        
        # Capturar a primeira página do dashboard
        screenshot_path = capturar_primeira_pagina()
        
        st.info("Enviando relatórios e captura de tela... Por favor, aguarde.")
        print("Iniciando processo de envio...")
        try:
            if enviar_whatsapp(numero_whatsapp, relatorio_principal):
                st.success("Relatório principal enviado com sucesso!")
                time.sleep(5)  # Espera 5 segundos entre os envios
                if enviar_whatsapp(numero_whatsapp, relatorio_ultimo_dia):
                    st.success("Relatório detalhado do último dia enviado com sucesso!")
                    time.sleep(5)  # Espera mais 5 segundos antes de enviar a imagem
                    if pywhatkit.sendwhats_image(numero_whatsapp, screenshot_path, "Captura da primeira página do dashboard"):
                        st.success("Captura de tela enviada com sucesso!")
                        time.sleep(5)  # Espera mais 5 segundos antes de fechar o navegador
                        fechar_navegador()
                        st.info("Navegador fechado. Processo concluído.")
                    else:
                        st.error("Erro ao enviar a captura de tela.")
                else:
                    st.error("Erro ao enviar o relatório detalhado do último dia.")
            else:
                st.error("Erro ao enviar o relatório principal.")
        except Exception as e:
            st.error(f"Erro ao enviar os relatórios: {str(e)}")
        print("Processo de envio concluído.")
