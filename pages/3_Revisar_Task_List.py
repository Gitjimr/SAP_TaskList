import numpy as np
import pandas as pd
import time
from unidecode import unidecode

import streamlit as st

st.session_state.update(st.session_state)
for k, v in st.session_state.items():
    st.session_state[k] = v

st.set_page_config(
    page_title="Revisão de Listas de Tarefas",
    # page_icon="🧊",
    layout="wide",
    # initial_sidebar_state="expanded",
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


# Função para realizar checagem: Exluir linhas completamente nulas, duplicatas e resetar o índice

def checagem_df(df):
    df = df.drop_duplicates()
    # df = df.dropna(how='all')
    df = df.reset_index(drop=True)

    return df

#

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
            SAP_MATERIAIS.dropna(subset='Material',inplace=True)
            SAP_MATERIAIS.reset_index(drop=True, inplace=True)

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


    st.header("Revisar Listas de Tarefas SAP", anchor=False)
    uploaded_file2 = st.file_uploader("Carregar Tabela de Listas de Tarefas",
                                             help="Carregar arquivo com as listas de tarefas no padrão correto para serem revisadas e associadas aos equipamentos e revisadas."
                                             )

    # Remover linhas com valores NaN na coluna 'Local de instalação'
    SAP_EQP_clean = SAP_EQP.dropna(subset=['Local de instalação'])

    # Aplicar a máscara booleana na coluna 'Local de instalação'
    SAP_EQP_filt = SAP_EQP_clean[SAP_EQP_clean['Local de instalação'].str.contains('-IND-')]

    lista_plantas = list(SAP_EQP_filt['Centro planejamento'].unique())
    lista_plantas.sort()

    plant_option = st.selectbox(label='Selecione a planta',
                                  options=lista_plantas
                                  )

    SAP_EQP_filt = SAP_EQP_filt[SAP_EQP_clean['Local de instalação'].str.contains(str(plant_option))]


    if uploaded_file2 is not None:
        #   Carregando arquivo op_padrao
        with st.spinner('Carregando Op_Padrao...'):
            try:
                skiprow = 0
                while 1:
                    OP_PADRAO = pd.read_excel(uploaded_file2, sheet_name="op_padrao", skiprows=skiprow, dtype=str)
                    lista_colunas_df_op = OP_PADRAO.columns.tolist()
                    if 'TEXTO TIPO EQUIPAMENTO' in lista_colunas_df_op:  # Identificar se é o cabeçalho
                        break
                    if skiprow > 10:
                        print('ERRO: REVISAR POSIÇÃO DO CABEÇALHO DA TABELA')
                        break
                    skiprow += 1
            except:
                st.write('ERRO AO CARREGAR OP_PADRAO')

        with st.spinner('Carregando EQUIPAMENTOS...'):
            try:
                EQUIPAMENTOS = pd.read_excel(uploaded_file2, sheet_name="EQUIPAMENTOS", skiprows=0, dtype=str)
                tem_planilha_eqp = 1
            except:
                tem_planilha_eqp = 0


        #   Ajustando OP_PADRAO

        OP_PADRAO.drop_duplicates(inplace=True)  # Para remover todas as linhas duplicadas
        OP_PADRAO.sort_values(by=['TEXTO TIPO EQUIPAMENTO'],
                              inplace=True)  # Ordem com index ordenado com base na task list
        OP_PADRAO.reset_index(drop=True, inplace=True)

        st.divider()
        #

        #   Revisão da op_padrao
        st.header('Realize as modificações da tabela abaixo e insira-a novamente.', divider='gray', anchor=False)

        colunas_op_padrao = [
            "TEXTO TIPO EQUIPAMENTO",
            "TIPO EQUIPAMENTO",
            "VARIACAO DESC",
            "VARIACAO N4/N5",
            "TASK LIST_TRECHO 01",
            "TIPO ATIV",
            "ROTA?",
            "TRECHO 02_TASK LIST",
            "ESPECIALIDADE",
            "PERIODICIDADE",
            "ESTADO MAQ",
            "TASK LIST_PARCIAL",
            "OPERACAO PADRAO",
            "TEXTO DESCRITIVO",
            "CT PM",
            "DUR_NORMAL_MIN",
            "QTD",
            "TRABALHO_MIN"
        ]

        if not all(coluna in OP_PADRAO.columns for coluna in colunas_op_padrao):
            st.write('ERRO: COLUNAS FALTANDO. RECOMENDA-SE VERIFICAR PADRÃO NA ABA DE CRIAÇÃO DE TASK LIST.')

        SAP_CTPM_filt = SAP_CTPM[SAP_CTPM['Cen.'] == str(plant_option)]
        lista_ctpms = SAP_CTPM_filt['CenTrab'].unique().tolist()

        lista_tipo_ativ = ['REVI', 'INSP', 'LUBR', 'TROC', 'LIMP', 'REAP', 'TEST', 'AJUS',
                 'CALI']

        lista_periodicidades = ['1D', '7D', '15D', '1M', '2M', '3M', '4M', '5M', '6M', '7M', '8M', '9M',
                                               '10M', '11M'
                                          , '1A', '2A', '3A', '4A', '5A', '6A', '7A', '8A', '9A', '10A',
                                               "1000H",
                                               "2000H",
                                               "3000H",
                                               "4000H",
                                               "5000H",
                                               "6000H",
                                               "7500H",
                                               "8000H",
                                               "10000H",
                                               "12000H",
                                               "15000H",
                                               "16000H",
                                               "20000H",
                                               "24000H",
                                               "25000H",
                                               "30000H",
                                               "50000H",
                                               "100000H"]

        lista_colunas_obrigatorias = ['TEXTO TIPO EQUIPAMENTO', 'TASK LIST_TRECHO 01', 'TIPO ATIV', 'ROTA?', 'ESPECIALIDADE', 'PERIODICIDADE', 'ESTADO MAQ', 'TASK LIST_PARCIAL', 'OPERACAO PADRAO', 'CT PM', 'DUR_NORMAL_MIN']

        tb_rev = {'ÍNDICE':[],'REVISÃO':[]}

        for i in range(len(OP_PADRAO['TEXTO TIPO EQUIPAMENTO'])):

            ##   Reatribuindo colunas de fórmula:
            try:
                if str(OP_PADRAO['ESTADO MAQ'][i]) != '1':
                    OP_PADRAO['TASK LIST_TRECHO 01'][i] = OP_PADRAO['TIPO ATIV'][i][0:4]+' '+OP_PADRAO['ESPECIALIDADE'][i][0:3]+' '+OP_PADRAO['PERIODICIDADE'][i]
                else:
                    OP_PADRAO['TASK LIST_TRECHO 01'][i] = OP_PADRAO['TIPO ATIV'][i][0:4] + ' ' + OP_PADRAO['ESPECIALIDADE'][
                        i][0:3] + ' ' +'FUNC '+ OP_PADRAO['PERIODICIDADE'][i]
            except:
                tb_rev['ÍNDICE'].append(i)
                tb_rev['REVISÃO'].append('HÁ ALGUMA COLUNA OBRIGATÓRIA VAZIA OU FORA DO PADRÃO.')

            try:
                if str(OP_PADRAO['ROTA?'][i]) != 'SIM':
                    OP_PADRAO['TASK LIST_PARCIAL'][i] = OP_PADRAO['TASK LIST_TRECHO 01'][i] + ' ' + OP_PADRAO['TRECHO 02_TASK LIST'][i]
            except:
                tb_rev['ÍNDICE'].append(i)
                tb_rev['REVISÃO'].append('ERRO COLUNA ROTA.')

            try:
                if int(OP_PADRAO['QTD'][i]) < 1:
                    OP_PADRAO['QTD'][i] = 1
                else:
                    OP_PADRAO['QTD'][i] = int(OP_PADRAO['QTD'][i])
            except:
                OP_PADRAO['QTD'][i] = 1

            try:
                OP_PADRAO['DUR_NORMAL_MIN'][i] = int(OP_PADRAO['DUR_NORMAL_MIN'][i])
                OP_PADRAO['TRABALHO_MIN'][i] = OP_PADRAO['DUR_NORMAL_MIN'][i] * OP_PADRAO['QTD'][i]
            except:
                tb_rev['ÍNDICE'].append(i)
                tb_rev['REVISÃO'].append('ERRO COLUNA DUR_NORMAL_MIN.')

            OP_PADRAO['VARIACAO N4/N5'][i] = '*'

            ##  Revisões:
            try:
                for coluna in lista_colunas_obrigatorias:
                    if pd.isna(OP_PADRAO[coluna][i]) or OP_PADRAO[coluna][i] == '':
                        tb_rev['ÍNDICE'].append(i)
                        tb_rev['REVISÃO'].append('HÁ ALGUMA COLUNA OBRIGATÓRIA VAZIA.')
            except:
                tb_rev['ÍNDICE'].append(i)
                tb_rev['REVISÃO'].append('HÁ ALGUMA COLUNA OBRIGATÓRIA VAZIA.')

            try:
                if not any(ativ in str(OP_PADRAO['TIPO ATIV'][i]) for ativ in lista_tipo_ativ):
                    tb_rev['ÍNDICE'].append(i)
                    tb_rev['REVISÃO'].append('TIPO DE ATIVIDADE NÃO REGISTRADO NO SAP.')
            except:
                tb_rev['ÍNDICE'].append(i)
                tb_rev['REVISÃO'].append('TIPO DE ATIVIDADE NÃO REGISTRADO NO SAP.')

            try:
                if OP_PADRAO['ROTA?'][i] != 'SIM' and (not isinstance(OP_PADRAO['TRECHO 02_TASK LIST'][i], str) or OP_PADRAO['TRECHO 02_TASK LIST'][i] == ''):
                    tb_rev['ÍNDICE'].append(i)
                    tb_rev['REVISÃO'].append('FALTOU ESPECIFICAR TEXTO NA COLUNA TRECHO 02_TASK LIST CASO NÃO SEJA ROTA (EQUIPAMENTO) OU ROTA PERSONALIZADA.')
            except:
                tb_rev['ÍNDICE'].append(i)
                tb_rev['REVISÃO'].append(
                    'FALTOU ESPECIFICAR TEXTO NA COLUNA TRECHO 02_TASK LIST CASO NÃO SEJA ROTA (EQUIPAMENTO) OU ROTA PERSONALIZADA.')

            try:
                if 'MEC' not in OP_PADRAO['ESPECIALIDADE'][i] and 'ELE' not in OP_PADRAO['ESPECIALIDADE'][i]:
                    tb_rev['ÍNDICE'].append(i)
                    tb_rev['REVISÃO'].append('ESPECIALIDADE DA LISTA DE TAREFAS DEVE SER MECANICA OU ELETRICA.')
            except:
                tb_rev['ÍNDICE'].append(i)
                tb_rev['REVISÃO'].append('ESPECIALIDADE DA LISTA DE TAREFAS DEVE SER MECANICA OU ELETRICA.')

            try:
                if not any(peri in str(OP_PADRAO['PERIODICIDADE'][i]) for peri in lista_periodicidades):
                    tb_rev['ÍNDICE'].append(i)
                    tb_rev['REVISÃO'].append('AJUSTAR PERIODICIDADE.')
            except:
                tb_rev['ÍNDICE'].append(i)
                tb_rev['REVISÃO'].append('AJUSTAR PERIODICIDADE.')

            try:
                if '1' not in str(OP_PADRAO['ESTADO MAQ'][i]) and '0' not in str(OP_PADRAO['ESTADO MAQ'][i]):
                    tb_rev['ÍNDICE'].append(i)
                    tb_rev['REVISÃO'].append('ESTADO DE MÁQUINA DEVE SER 1 (FUNCIONAMENTO) OU 0 (PARADA).')
            except:
                tb_rev['ÍNDICE'].append(i)
                tb_rev['REVISÃO'].append('ESTADO DE MÁQUINA DEVE SER 1 (FUNCIONAMENTO) OU 0 (PARADA).')

            try:
                if OP_PADRAO['CT PM'][i] not in lista_ctpms:
                    tb_rev['ÍNDICE'].append(i)
                    tb_rev['REVISÃO'].append('CTPM NÃO EXISTE NA UNIDADE SELECIONADA.')
            except:
                tb_rev['ÍNDICE'].append(i)
                tb_rev['REVISÃO'].append('CTPM NÃO EXISTE NA UNIDADE SELECIONADA.')

            try:
                if len(str(OP_PADRAO['TASK LIST_PARCIAL'][i])) > 35:
                    tb_rev['ÍNDICE'].append(i)
                    tb_rev['REVISÃO'].append('TRECHO 02 TASK LIST: QUANTIDADE DE CARACTERES ULTRAPASSANDO O MÁXIMO.')
            except:
                tb_rev['ÍNDICE'].append(i)
                tb_rev['REVISÃO'].append('ERRO NA COLUNA TASK LIST_PARCIAL.')

            try:
                if len(str(OP_PADRAO['OPERACAO PADRAO'][i])) > 40:
                    tb_rev['ÍNDICE'].append(i)
                    tb_rev['REVISÃO'].append('OPERACAO PADRÃO: QUANTIDADE DE CARACTERES ULTRAPASSANDO O MÁXIMO.')
            except:
                tb_rev['ÍNDICE'].append(i)
                tb_rev['REVISÃO'].append('ERRO NA COLUNA DE OPERACAO PADRAO.')

            if isinstance(OP_PADRAO['OPERACAO PADRAO'][i], str):
                if not any(name in OP_PADRAO['OPERACAO PADRAO'][i] for name in
                           ['INSP', 'TROC', 'REV', 'BLOQ', 'DESBLOQ', 'DOSA', 'SUB', 'VERI', 'DESL', 'LIG', 'CONEC',
                            'DESCON', 'REAL', 'LIMP', 'EFET', 'EXEC', 'TEST', 'ISOL', 'RETI', 'ABR', 'FECH', 'LUB',
                            'ESGOT',
                            'SOLIC', 'ENSA', 'CALI', 'AJUS', 'REAP', 'FIX', 'PARAF', 'INSER', 'COLOC',
                            'AJUS']) and not any(
                    letr_2 in OP_PADRAO['OPERACAO PADRAO'][i].split()[0][-2:] for letr_2 in ['AR', 'ER', 'IR']):
                    tb_rev['ÍNDICE'].append(i)
                    tb_rev['REVISÃO'].append('OPERACAO PADRAO: VERIFICAR ORTOGRAFIA OU SE HÁ A AÇÃO DESEMPENHADA (INSP, REVI, ETC). CASO ERRO PERSISTA, COLOQUE UM VERBO NO INÍCIO.')

        #

        st.dataframe(
            OP_PADRAO,
            column_order=[
                'TEXTO TIPO EQUIPAMENTO', 'TIPO ATIV', 'ROTA?', 'TRECHO 02_TASK LIST',
                'ESPECIALIDADE','PERIODICIDADE','ESTADO MAQ','TASK LIST_PARCIAL','OPERACAO PADRAO','TEXTO DESCRITIVO',
                'CT PM', 'DUR_NORMAL_MIN',
                'QTD', 'TRABALHO_MIN'],
        )

        col1, col2 = st.columns([4, 2])

        col1.subheader('Revisões necessárias por linha da tabela', divider='gray',anchor=False)

        df_tb_rev = pd.DataFrame(tb_rev)

        col1.dataframe(data=df_tb_rev, hide_index=True, use_container_width=True)

        # Download Button
        col2.subheader("⬇️ Download", divider='gray', anchor=False)

        col2.markdown("""❗  Após a conclusão: enviar arquivo para :red[João Ivo] (ter08334@mdb.com.br)
                          """)

        try:
            ## Filtrar Planilha EQP
            ### Lista das três primeiras colunas do df1 para verificar a presença dos valores de df2
            colunas_verificadas_EQP = ['Tipo OP Padrão N6', 'Tipo OP Padrão N7', 'Tipo OP Padrão N8']

            ### Filtrando as linhas de df1
            EQUIPAMENTOS = EQUIPAMENTOS[EQUIPAMENTOS[colunas_verificadas_EQP].apply(lambda row: any(row.isin(OP_PADRAO['TEXTO TIPO EQUIPAMENTO'])), axis=1)]
        except:
            pass

        import io

        buffer3 = io.BytesIO()
        with pd.ExcelWriter(buffer3, engine="xlsxwriter") as writer:
            try:
                EQUIPAMENTOS.to_excel(writer, sheet_name="EQUIPAMENTOS", index=False)
            except:
                pass
            OP_PADRAO.to_excel(writer, sheet_name="op_padrao", index=False)
            try:
                df_tb_rev.to_excel(writer, sheet_name="REVISÕES", index=False)
            except:
                pass
            # Close the Pandas Excel writer and output the Excel file to the buffer
            writer.close()

            col2.download_button(
                label="Download Listas de Tarefas Revisada",
                data=buffer3,
                file_name="Op_Padrao_"+str(plant_option)+".xlsx",
                use_container_width = True,
                #disabled=not df_tb_rev.empty

            )

