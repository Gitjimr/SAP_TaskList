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

st.markdown(
    """
    <style>
    .st-emotion-cache-ch5dnh
    {
        visibility: hidden;
    }
    .st-emotion-cache-q16mip
    {
        visibility: hidden;
    }
    .st-emotion-cache-ztfqz8
    {
        visibility: hidden;
    }
    </style>
    """,
    unsafe_allow_html=True
)

prcnt_width = 80
max_width_str = f"max-width: {prcnt_width}%;"
st.markdown(f""" 
            <style> 
            .reportview-container .main .block-container{{{max_width_str}}}
            </style>    
            """,
            unsafe_allow_html=True)

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


    st.header('Identificar Equipamentos Principais com Planos Carregados', divider='gray', anchor=False)

    if 'DF_PLAN_EQP' not in st.session_state:
        with st.spinner("Carregando..."):

            # Leitura dos arquivos baixados

            ia39 = SAP_TL
            ip18 = SAP_ITEM

            ia39 = ia39[['Equipamento operação', 'Descrição','Grupo']].copy()
            ia39.rename(columns={'Equipamento operação': 'ID Equipamento'}, inplace = True)
            ip18 = ip18[['Plano de manutenção', 'Item manutenção', 'Grupo', 'Equipamento','Descrição item de manutenção']].copy()
            ip18 = ip18[['Equipamento','Descrição item de manutenção','Grupo','Plano de manutenção','Item manutenção']].copy()
            ip18.rename(columns={'Equipamento': 'ID Equipamento'}, inplace = True)
            ip18.rename(columns={'Descrição item de manutenção': 'Descrição'}, inplace = True)


            #   SEPARAR EQUIPAMENTOS, LISTAS E PLANOS

            # Criar uma coluna 'Plano de manutenção' no DataFrame ia39 preenchida com valores NaN
            ia39['Plano de manutenção'] = np.nan
            ia39['Item manutenção'] = np.nan

            # Criar um dicionário a partir dos valores de 'Grupo' como chave e os valores correspondentes de 'Plano de manutenção' como valor
            grupo_plano_dict = ip18.set_index('Grupo')['Plano de manutenção'].to_dict()
            grupo_item_dict = ip18.set_index('Grupo')['Item manutenção'].to_dict()

            # Mapear os valores de 'Grupo' de ia39 para os valores correspondentes de 'Plano de manutenção' usando o dicionário
            ia39['Plano de manutenção'] = ia39['Grupo'].map(grupo_plano_dict)
            ia39['Item manutenção'] = ia39['Grupo'].map(grupo_item_dict)

            # Caso não encontrado
            ia39['Plano de manutenção'] = ia39['Plano de manutenção'].fillna('SEM PLANO')
            ia39['Item manutenção'] = ia39['Item manutenção'].fillna('SEM ITEM')


            eqp_tl_plan = pd.concat([ip18, ia39], ignore_index=True, sort=False)
            eqp_tl_plan.drop_duplicates()
            eqp_tl_plan.dropna(subset=['ID Equipamento'], inplace=True)
            eqp_tl_plan.reset_index(drop=True, inplace=True)


            # Removendo equipamentos não principais do eqp_tl_plan
            equipamentos_principais = SAP_EQP_N6['ID_SAP_N6']
            eqp_tl_plan = eqp_tl_plan[eqp_tl_plan['ID Equipamento'].isin(equipamentos_principais)]

            # Identificando equipamentos principais que não estão no eqp_tl_plan
            equipamentos_principais_faltantes = SAP_EQP_N6['ID_SAP_N6']
            equipamentos_principais_faltantes = equipamentos_principais_faltantes[~equipamentos_principais_faltantes.isin(eqp_tl_plan['ID Equipamento'])]

            # Novo DataFrame com as linhas a serem adicionadas
            novas_linhas = pd.DataFrame({
                'ID Equipamento': equipamentos_principais_faltantes,
                'Descrição': 'N/A',
                'Grupo': 'SEM LISTA',
                'Plano de manutenção': 'SEM PLANO'
            })

            # Concatene os DataFrames
            eqp_tl_plan = pd.concat([eqp_tl_plan, novas_linhas], ignore_index=True)


            # Dicionário mapeando os valores de 'ID Equipamento' para 'ID_SAP_N6'
            id_equipamento_to_id_sap_n6 = SAP_EQP_N6.set_index('ID_SAP_N6')['LI_N5'].to_dict()

            # Mapear 'ID Equipamento' para 'ID_SAP_N6' e, em seguida, associe os valores correspondentes de 'LI_N5'
            #eqp_tl_plan_com_li = eqp_tl_plan.copy()  # Copie o DataFrame principal para manter sua integridade
            eqp_tl_plan['LI_N5'] = eqp_tl_plan['ID Equipamento'].map(id_equipamento_to_id_sap_n6)

            # Fazendo o mesmo para a desc do equip
            id_equipamento_to_id_sap_n6 = SAP_EQP_N6.set_index('ID_SAP_N6')['EQUIPAMENTO PRINCIPAL'].to_dict()
            eqp_tl_plan['EQUIPAMENTO PRINCIPAL'] = eqp_tl_plan['ID Equipamento'].map(id_equipamento_to_id_sap_n6)
            column_to_move = eqp_tl_plan.pop("EQUIPAMENTO PRINCIPAL")
            eqp_tl_plan.insert(1, "EQUIPAMENTO PRINCIPAL", column_to_move)

            #  Trazendo últimas ordens e datas
            ## Filtrar as entradas onde a coluna Order não está vazia
            SAP_PMI_filt = SAP_PMI.dropna(subset=['Order'])
            ## Agrupar por item e pegar a última entrada em cada grupo
            data_ultima_por_item = SAP_PMI_filt.groupby('Maintenance Item')['Scheduled start date'].last()
            ordem_ultima_por_item = SAP_PMI_filt.groupby('Maintenance Item')['Order'].last()
            ## Mapear 'ID_SAP_N6' para o último 'EQUIPAMENTO PRINCIPAL'
            data_ultima_por_item = data_ultima_por_item.to_dict()
            ordem_ultima_por_item = ordem_ultima_por_item.to_dict()
            ## Mapear 'ID Equipamento' para o último 'EQUIPAMENTO PRINCIPAL'
            eqp_tl_plan['ULTIMA DATA'] = eqp_tl_plan['Item manutenção'].map(data_ultima_por_item)
            eqp_tl_plan['ULTIMA ORDEM'] = eqp_tl_plan['Item manutenção'].map(ordem_ultima_por_item)


            #   CRIAR E FORMATAR DF COM BASE NA BASE DE EQUIPAMENTOS

            df_plan_eqp = eqp_tl_plan

            df_plan_eqp['GRUPO'] = df_plan_eqp['LI_N5'].apply(lambda x: str(x)[9:12])
            df_plan_eqp['PLANTA'] = df_plan_eqp['LI_N5'].apply(lambda x: str(x)[0:4])

            df_plan_eqp['Plano de manutenção'] = df_plan_eqp['Plano de manutenção'].fillna('SEM ID')
            df_plan_eqp['Grupo'] = df_plan_eqp['Grupo'].fillna('SEM ID')
            df_plan_eqp['LI_N5'] = df_plan_eqp['LI_N5'].fillna('SEM LI')

            df_plan_eqp['TEM PLANO/LISTA?'] = df_plan_eqp['Plano de manutenção'].apply(lambda x: 'SEM PLANO' if x == "SEM PLANO" else "CONTÉM")
            df_plan_eqp['TEM PLANO/LISTA?'] = df_plan_eqp['Grupo'].apply(lambda x: 'SEM LISTA' if x == "SEM LISTA" else "CONTÉM")

            df_plan_eqp = df_plan_eqp[df_plan_eqp['LI_N5'].str.contains('-IND-')]
            df_plan_eqp.drop_duplicates(subset=None, inplace=True)  # Para remover todas as linhas duplicadas
            df_plan_eqp.sort_values(by=['ID Equipamento','TEM PLANO/LISTA?'], inplace=True)      # Ordem com index ordenado com base na task list
            df_plan_eqp.reset_index(drop=True, inplace=True)

            #if 'DF_PLAN_EQP' not in st.session_state:
            st.session_state.DF_PLAN_EQP = df_plan_eqp
    else:
        df_plan_eqp = st.session_state['DF_PLAN_EQP']

    col1, col2 = st.columns([4, 2])

    #   SETUP PLANTA E GRUPO

    options_planta = col1.multiselect(
        label='Selecione a(s) planta(s)',
        options=sorted(list(df_plan_eqp['PLANTA'].unique())),
        key='OP_PLANTA'
        )

    try:
        grupos_por_planta = df_plan_eqp[df_plan_eqp['PLANTA'].isin(options_planta)]['GRUPO'].unique()
        options_grupo = col1.multiselect(
            label='Selecione o(s) grupo(s)',
            disabled = True if len(grupos_por_planta)==0 else False,
            options=sorted(list(grupos_por_planta)) if len(grupos_por_planta)>0 else [],
            key='OP_GRUPO' if len(grupos_por_planta)>0 else None
        )
    except:
        del st.session_state['OP_GRUPO']
        options_grupo = col1.multiselect(
            label='Selecione o(s) grupo(s)',
            disabled=True if len(grupos_por_planta) == 0 else False,
            options=[]
        )

    options_tem_plano = col1.multiselect(
        label='Status plano',
        options=['CONTÉM', 'SEM PLANO', 'SEM LISTA'],
        key='OP_TEMPLANO'
        )

    if 'OP_PLANTA' not in st.session_state:
        st.session_state.OP_PLANTA = options_planta
    if 'OP_GRUPO' not in st.session_state:
        st.session_state.OP_GRUPO = options_grupo
    if 'OP_TEMPLANO' not in st.session_state:
        st.session_state.OP_TEMPLANO = options_tem_plano


    #   GRÁFICOS

    ## FAZER HEATMAP:

    df_heatmap = pd.DataFrame({'PLANTA': [], '% CARREGADO': []})
    for fabrica in list(df_plan_eqp['PLANTA'].unique()):
        try:
            count_tot = (df_plan_eqp['PLANTA'] == str(fabrica)).sum()
            count_templano = ((df_plan_eqp['TEM PLANO/LISTA?'] == 'CONTÉM') & (df_plan_eqp['PLANTA'] == str(fabrica))).sum()
            df_heatmap = pd.concat(
                [df_heatmap, pd.DataFrame({'PLANTA': [str(fabrica)], '% CARREGADO': [count_templano / count_tot]})],
                ignore_index=True, sort=False)
            df_heatmap.sort_values(by=['PLANTA'], inplace=True)
            df_heatmap.reset_index(drop=True,inplace=True)
        except:
            pass

    col2.subheader('% Equipamentos com Planos Carregados por Fábrica',anchor=False)
    col2.table(df_heatmap.style.background_gradient(subset=["% CARREGADO"], cmap="RdYlGn", vmin=0, vmax=1))


    ##   DISPLAY DE DATAFRAME COM FILTROS:

    if options_grupo is not None and options_planta is not None and options_tem_plano is not None:
        with st.spinner('Carregando Tabela...'):
            #df_plan_eqp_option = df_plan_eqp[df_plan_eqp['PLANTA'].isin(options_planta) & df_plan_eqp['GRUPO'].isin(options_grupo) & df_plan_eqp['Plano de manutenção'].isin(options_tem_plano) & df_plan_eqp['Grupo'].isin(options_tem_plano)]
            df_plan_eqp_option = df_plan_eqp[
                (df_plan_eqp['PLANTA'].isin(options_planta)) &
                (df_plan_eqp['GRUPO'].isin(options_grupo)) &
                ((df_plan_eqp['Plano de manutenção'].isin(options_tem_plano)) |
                (df_plan_eqp['Grupo'].isin(options_tem_plano)) |
                 df_plan_eqp['TEM PLANO/LISTA?'].isin(options_tem_plano))
                ]
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

                col1.subheader("⬇️ Download", divider='gray', anchor=False)
                col1.download_button(
                    label="Download Relação PLANO/TL x EQP Filtrada",
                    data=buffer1,
                    file_name='PLANO-TLxEQP' + '.xlsx',
                )
                ###




