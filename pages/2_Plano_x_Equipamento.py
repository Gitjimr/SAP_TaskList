import streamlit as st
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import plotly.graph_objects as go
from PIL import Image
import os

st.session_state.update(st.session_state)
for k, v in st.session_state.items():
    st.session_state[k] = v

st.set_page_config(
    page_title='% Equipamento x Plano',
    layout='wide'
)

hide_menu = '''
        <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        </style>
'''
st.markdown(hide_menu, unsafe_allow_html=True)

st.markdown(
    """
    <style>
    .css-1jc7ptx, .e1ewe7hr3, .viewerBadge_container__1QSob,
    .styles_viewerBadge__1yB5_, .viewerBadge_link__1S137,
    .viewerBadge_text__1JaDK {
        display: none;
    }
    </style>
    """,
    unsafe_allow_html=True
)



import os

path = os.path.dirname(__file__)
my_path = path + '/files/'

uploaded_file = st.sidebar.file_uploader("Carregar Dados Chave",
help="Carregar arquivo com dados necessários do SAP. Caso precise recarregá-lo, atualize a página. Este arquivo deve ser continuamente atualizado conforme novos dados sejam inseridos no SAP"
                                         )

if uploaded_file is not None or 'SAP_CTPM' in st.session_state:
    if 'SAP_CTPM' not in st.session_state:
        with st.spinner('Carregando Lista de Equipamentos...'):
            SAP_EQP_N6 = pd.read_excel(uploaded_file, sheet_name="EQP", skiprows=0,dtype=str)
        with st.spinner('Carregando IE03 SAP...'):
            SAP_EQP = pd.read_excel(uploaded_file, sheet_name="IE03", skiprows=0, dtype=str)
        with st.spinner('Carregando IA39 SAP...'):
            SAP_TL = pd.read_excel(uploaded_file, sheet_name="IA39", skiprows=0, dtype=str)
        with st.spinner('Carregando IP18 SAP...'):
            SAP_ITEM = pd.read_excel(uploaded_file, sheet_name="IP18", skiprows=0, dtype=str)
        with st.spinner('Carregando IP24 SAP...'):
            SAP_PMI = pd.read_excel(uploaded_file, sheet_name="IP24", skiprows=0, dtype=str)
        with st.spinner('Carregando Centros de Trabalho SAP...'):
            SAP_CTPM = pd.read_excel(uploaded_file, sheet_name="CTPM", skiprows=0, dtype=str)
        with st.spinner('Carregando Materiais SAP...'):
            SAP_MATERIAIS = pd.read_excel(uploaded_file, sheet_name="MATERIAIS", skiprows=0, dtype=str)

            st.session_state.SAP_EQP_N6 = SAP_EQP_N6
            st.session_state.SAP_EQP = SAP_EQP
            st.session_state.SAP_TL = SAP_TL
            st.session_state.SAP_ITEM = SAP_ITEM
            st.session_state.SAP_PMI = SAP_PMI
            st.session_state.SAP_MATERIAIS = SAP_MATERIAIS
            st.session_state.SAP_CTPM = SAP_CTPM
    else:
        SAP_EQP_N6 = st.session_state['SAP_EQP_N6']
        SAP_EQP = st.session_state['SAP_EQP']
        SAP_TL = st.session_state['SAP_TL']
        SAP_ITEM = st.session_state['SAP_ITEM']
        SAP_PMI = st.session_state['SAP_PMI']
        SAP_CTPM = st.session_state['SAP_CTPM']
        SAP_MATERIAIS = st.session_state['SAP_MATERIAIS']


    st.header('Identificar Equipamentos com Planos Carregados', divider='gray', anchor=False)


    # Leitura dos arquivos baixados

    ia39 = SAP_TL
    ip18 = SAP_ITEM


    ia39 = ia39[['Equipamento operação', 'Grupo']].copy()
    ia39.rename(columns={'Equipamento operação': 'EQP_N6'}, inplace = True)


    st.write(ia39)
    st.write(ip18)

    #   SEPARAR EQUIPAMENTOS, LISTAS E PLANOS

    # Criar uma coluna 'Plano de manutenção' no DataFrame ia39 preenchida com valores NaN
    ia39['Plano de manutenção'] = np.nan

    # Criar um dicionário a partir dos valores de 'Grupo' como chave e os valores correspondentes de 'Plano de manutenção' como valor
    grupo_plano_dict = ip18.set_index('Grupo')['Plano de manutenção'].to_dict()

    # Mapear os valores de 'Grupo' de ia39 para os valores correspondentes de 'Plano de manutenção' usando o dicionário
    ia39['Plano de manutenção'] = ia39['Grupo'].map(grupo_plano_dict)

    st.write(ia39)
    st.write(ip18)

    eqp_tl_plan = pd.concat([ip18, ia39], ignore_index=True, sort=False)
    eqp_tl_plan.drop_duplicates()
    eqp_tl_plan.dropna(subset=['EQP_N6'], inplace=True)
    eqp_tl_plan.reset_index(drop=True, inplace=True)

    st.write(eqp_tl_plan)


    #   CRIAR E FORMATAR DF COM BASE NA BASE DE EQUIPAMENTOS

    df_plan_eqp = SAP_EQP_N6
    #df_plan_eqp['TEM PLANO?'] = df_plan_eqp['EQUIPAMENTO PRINCIPAL'].apply(lambda x: 'SIM' if x in eqp_tl_plan['EQP_N6'].values else 'SEM PLANO')
    df_plan_eqp['TEM PLANO?'] = df_plan_eqp['ID_SAP_N6'].isin(eqp_tl_plan['EQP_N6']).map(
        {True: 'SIM', False: 'SEM PLANO'})
    df_plan_eqp['GRUPO'] = df_plan_eqp['LI_N5'].apply(lambda x: str(x)[9:12])
    df_plan_eqp['PLANTA'] = df_plan_eqp['LI_N5'].apply(lambda x: str(x)[0:4])

    df_plan_eqp['PLANO'] = np.nan
    df_plan_eqp['LT'] = np.nan


    #   TRAZER PARA CADA EQP OS PLANOS E LISTAS

    # Converter os valores np.nan para listas vazias
    eqp_tl_plan['Plano de manutenção'] = eqp_tl_plan['Plano de manutenção'].apply(
        lambda x: '' if isinstance(x, float) and np.isnan(x) else x)
    eqp_tl_plan['Grupo'] = eqp_tl_plan['Grupo'].apply(
        lambda x: '' if isinstance(x, float) and np.isnan(x) else x)

    # Agrupar e unir os planos de manutenção e grupos para cada equipamento em eqp_tl_plan
    eqp_tl_plan_grouped = eqp_tl_plan.groupby('EQP_N6').agg({
        'Plano de manutenção': lambda x: ','.join(map(str, x)),
        'Grupo': lambda x: ','.join(map(str, x))
    }).reset_index()

    # Juntar os DataFrames usando merge
    df_plan_eqp = pd.merge(df_plan_eqp, eqp_tl_plan_grouped, how='left', left_on='ID_SAP_N6', right_on='EQP_N6')

    # Preencher os valores nulos nas colunas 'PLANO' e 'LT' com '-'
    df_plan_eqp['PLANO'] = df_plan_eqp['Plano de manutenção'].fillna('-')
    df_plan_eqp['LT'] = df_plan_eqp['Grupo'].fillna('-')

    # Remover as colunas extras
    df_plan_eqp.drop(['Plano de manutenção', 'Grupo', 'EQP_N6'], axis=1, inplace=True)

    #st.write()

    df_plan_eqp['PLANO'].astype(str).apply(lambda x: ','.join(sorted(pd.unique(x.split(',')))))#*
    df_plan_eqp['LT'].astype(str).apply(lambda x: ','.join(sorted(pd.unique(x.split(',')))))#*

    #df_plan_eqp['PLANO'].astype(str).apply(lambda x: ','.join(sorted(set(x.replace('[]','').replace(',,',',').split(',')))))#*
    #df_plan_eqp['LT'].astype(str).apply(lambda x: ','.join(sorted(set(x.replace('[]','').replace(',,',',').split(',')))))#*

    # aaa

    st.write(df_plan_eqp)

    df_plan_eqp.drop_duplicates(subset='ID_SAP_N6')
    df_plan_eqp.dropna(how='all')
    df_plan_eqp.reset_index(drop=True)

    st.session_state.df_plan_eqp = df_plan_eqp

    df_plan_eqp = st.session_state['df_plan_eqp']

    col1, col2 = st.columns([4, 2])

    #   SETUP PLANTA E GRUPO

    options_planta = col1.multiselect(
        label='Selecione a(s) planta(s)',
        options=sorted(list(df_plan_eqp['PLANTA'].unique()))
        )

    options_grupo = col1.multiselect(
        label='Selecione o(s) grupo(s)',
        options=sorted(list(df_plan_eqp['GRUPO'].unique()))
        )

    options_tem_plano = col1.multiselect(
        label='Status plano',
        options=['SIM', 'SEM PLANO'],
        default='SEM PLANO'
        )

    #   GRÁFICOS

    ## FAZER HEATMAP:

    df_heatmap = pd.DataFrame({'PLANTA': [], '% CARREGADO': []})
    for fabrica in list(df_plan_eqp['PLANTA'].unique()):
        try:
            count_tot = (df_plan_eqp['PLANTA'] == str(fabrica)).sum()
            count_templano = ((df_plan_eqp['TEM PLANO?'] == 'SIM') & (df_plan_eqp['PLANTA'] == str(fabrica))).sum()
            df_heatmap = pd.concat(
                [df_heatmap, pd.DataFrame({'PLANTA': [str(fabrica)], '% CARREGADO': [count_templano / count_tot]})],
                ignore_index=True, sort=False)
        except:
            pass

    col2.subheader('% Equipamentos com Planos Carregados por Fábrica',anchor=False)
    col2.table(df_heatmap.style.background_gradient(subset=["% CARREGADO"], cmap="RdYlGn", vmin=0, vmax=1))


    ##   DISPLAY DE DATAFRAME COM FILTROS:

    if options_grupo is not None and options_planta is not None and options_tem_plano is not None:
        with st.spinner('Carregando Gráficos...'):
            df_plan_eqp_option = df_plan_eqp[df_plan_eqp['PLANTA'].isin(options_planta) & df_plan_eqp['GRUPO'].isin(options_grupo) & df_plan_eqp['TEM PLANO?'].isin(options_tem_plano)]

            df_plan_eqp_option.drop_duplicates(subset=['ID_SAP_N6'])
            df_plan_eqp_option.dropna(how='all')
            df_plan_eqp_option.reset_index(drop=True)

            col1.dataframe(df_plan_eqp_option)

            # Salvando em arquivo excel
            import io
            buffer1 = io.BytesIO()
            with pd.ExcelWriter(buffer1, engine="xlsxwriter") as excel_writer:
                ## Crie um objeto ExcelWriter
                nome_arquivo_sap = 'CHECK_PLXEQP'

                ## Salve cada DataFrame em uma planilha diferente
                df_plan_eqp_option.to_excel(excel_writer, sheet_name='PLANO-TLxEQP', index=False)

                ## Feche o objeto ExcelWriter
                excel_writer.close()

                col1.download_button(
                    label="Download Relação PLANO/TL x EQP Filtrada",
                    data=buffer1,
                    file_name='PLANO-TLxEQP' + '.xlsx',
                )
                ###




