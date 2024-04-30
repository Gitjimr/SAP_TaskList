import numpy as np
import pandas as pd
import time
from unidecode import unidecode

import streamlit as st
st.session_state.update(st.session_state)
for k, v in st.session_state.items():
    st.session_state[k] = v

st.set_page_config(
    page_title="Geração de Listas de Tarefas",
    #page_icon="🧊",
    layout="wide",
    #initial_sidebar_state="expanded",
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
    .st-emotion-cache-ch5dnh ef3psqc5
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
    df = df.dropna(how='all')
    df = df.reset_index(drop=True)

    return df


#

import os

path = os.path.dirname(__file__)
my_path = path + '/pages/files/'

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

    # Remover linhas com valores NaN na coluna 'Local de instalação'
    SAP_EQP_clean = SAP_EQP.dropna(subset=['Local de instalação'])

    # Aplicar a máscara booleana na coluna 'Local de instalação'
    SAP_EQP_filt = SAP_EQP_clean[SAP_EQP_clean['Local de instalação'].str.contains('-IND-')]

    lista_plantas = list(SAP_EQP_filt['Centro planejamento'].unique())
    lista_plantas.sort()

    st.header("Listas de Tarefas SAP")

    col1, col2 = st.columns([2, 4])

    plant_option = col1.selectbox(label='Selecione a planta',
                                options=lista_plantas,key='PLANT_OPTION'
    )

    SAP_EQP_filt = SAP_EQP_filt[SAP_EQP_clean['Local de instalação'].str.contains(str(plant_option))]


    #   Selecionar equipamentos
    try:
        op_eqp = col1.multiselect(
            label='Selecione o ID SAP dos equipamentos a receberem a operação',
            options=list(SAP_EQP_filt['Equipamento'].unique()),
            key='OP_EQP'
        )
    except:
        del st.session_state['OP_EQP']
        op_eqp = col1.multiselect(
            label='Selecione o ID SAP dos equipamentos a receberem a operação',
            options=list(SAP_EQP_filt['Equipamento'].unique())
        )
    op_eqp_ = []
    op_eqp_id = []
    if not isinstance(op_eqp, str):
        for option in op_eqp:
            index_procurado = SAP_EQP_filt[SAP_EQP_filt['Equipamento'] == option].index[0]
            op_eqp_.append(SAP_EQP_filt['Denominação do objeto técnico'][index_procurado])
            op_eqp_id.append(option)
            col1.write(SAP_EQP_filt['Denominação do objeto técnico'][index_procurado])
    else:
        index_procurado = SAP_EQP_filt[SAP_EQP_filt['Equipamento'] == op_eqp].index[0]
        op_eqp_.append(SAP_EQP_filt['Denominação do objeto técnico'][index_procurado])
        op_eqp_id.append(op_eqp)
        col1.write(SAP_EQP_filt['Denominação do objeto técnico'][index_procurado])

    col1.markdown("[Pesquisar Código SAP de Equipamento](#4aa479eb)", unsafe_allow_html=True)

    col1.divider()

    #   Selecionar operação
    op_oppadrao = col1.text_input(label='Operação Padrão',
                                max_chars=40,key='OP_OPPADRAO',help='Escreva o texto descritivo principal da operação em 40 caracteres. Abrevie se necessário.')
    op_optextolongo = col1.text_area(label='Texto Longo',key='OP_OPTEXTOLONGO',help='Escreva um texto detalhado da operação ou um que contenha detalhes adicionais não presentes no texto principal.')
    op_oppadrao = unidecode(op_oppadrao.upper())
    op_oppadrao = op_oppadrao.strip()
    op_optextolongo = unidecode(op_optextolongo.upper())
    op_optextolongo = op_optextolongo.strip()

    #   Selecionar centro de trabalho
    SAP_CTPM_filt = SAP_CTPM[SAP_CTPM['Cen.'] == str(plant_option)]
    lista_ctpms = SAP_CTPM_filt['CenTrab'].unique().tolist()
    try:
        op_ctpm = col1.selectbox(
            label='Centro de Trabalho',
            options=lista_ctpms,
            key='OP_CTPM',
        )
    except:
        if 'OP_CTPM' in st.session_state:
            del st.session_state['OP_CTPM']
            op_ctpm = col1.selectbox(
                label='Centro de Trabalho',
                options=lista_ctpms,
            )
    index_procurado = SAP_CTPM_filt[SAP_CTPM_filt['CenTrab'] == op_ctpm].index[0]
    col1.write(SAP_CTPM_filt['Descrição breve'][index_procurado])

    col1.divider()

    #   Selecionar tipo de ativ
    op_tipoativ = col1.selectbox(
        label='Tipo de Atividade',
        options=['REVISÃO','INSPEÇÃO','LUBRIFICAÇÃO','TROCA','LIMPEZA','REAPERTO','TESTE','AJUSTE','CALIBRAÇÃO'],key='OP_TIPOATIV'
    )

    #   Selecionar periodicidade
    op_periodicidade = col1.selectbox(label='Periodicidade',
    options=['1D','7D','15D','1M','2M','3M','4M','5M','6M','7M','8M','9M','10M','11M'
        ,'1A','2A','3A','4A','5A','6A','7A','8A','9A','10A',
        "1000H",
        "2000H",
        "4000H",
        "5000H",
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
        "100000H"],
        help='Exemplo: Mensal = 1M.',key='OP_PERIODICIDADE'
        )
    op_periodicidade = op_periodicidade.upper()

    #   Selecionar estado máq.
    op_estadomaq = col1.radio(
        label='Estado de Máquina',
        options=[0, 1],key='OP_ESTADOMAQ',help='0: Parada | 1: Em funcionamento'
    )
    op_estadomaq_ = 'FUNC' if op_estadomaq==1 else ''

    #   Selecionar se é rota e complemento da task list
    op_erota = col1.radio(
        label='Rota?',
        options=['SIM', 'NÃO'],key='OP_EROTA'
    )
    if op_erota == 'NÃO':
        op_trecho2tasklist = col1.text_input(label='Descrição de equipamento/componente da Lista de Tarefa',
                            max_chars=20-len(op_periodicidade) if op_estadomaq == 1 else 25-len(op_periodicidade),# if op_erota == 'NÃO' and op_estadomaq == 0 else '',
                            help='Texto que aparecerá no cabeçalho da lista de tarefas e do item no SAP. Caso seja rota, será o local de instalação.',key='OP_TRECHO2TASKLIST')
        op_trecho2tasklist = op_trecho2tasklist.upper()
        op_trecho2tasklist = op_trecho2tasklist.strip()
    else:
        op_trecho2tasklist = ''

    #   Selecionar especialidade da op
    op_espec = col1.radio(
        label='Tipo de Operação',
        options=['MECANICA', 'ELETRICA'],key='OP_ESPEC'
    )

    col1.divider()

    #   Selecionar duração e trabalho
    op_duracao = col1.number_input(label='Duração (Minutos)', format='%d', step=1, min_value=1,key='OP_DURACAO')
    op_operad = col1.number_input(label='Quantidade de operadores', format='%d', step=1,
                                min_value=1,help='Quantidade de pessoas executando esta operação.',key='OP_OPERAD')
    op_trabtot = op_duracao*op_operad

    st.divider()

    # Selecionar Material
    col1.divider()

    op_temmaterial = col1.radio(
        label='Operação possui material?',
        options=['NÃO', 'SIM'],key='OP_TEMMATERIAL'
    )
    if op_temmaterial == 'SIM':
        try:
            op_material = col1.selectbox(label='Selecione o código do material',
                                options=list(SAP_MATERIAIS['Material']),
                                key='OP_MATERIAL'
                                         )

            index_procurado = SAP_MATERIAIS[SAP_MATERIAIS['Material'] == op_material].index[0]
            op_undmaterial = SAP_MATERIAIS['UM básica'][index_procurado]
            col1.write(SAP_MATERIAIS['Texto breve material'][index_procurado])

            op_qtdmaterial = col1.number_input(label='Quantidade do material',
                                format='%f',
                                step=1.,
                                key='OP_QTDMATERIAL'
                                         )
        except:
            op_material = col1.selectbox(label='Selecione o código do material',
                                         options=list(SAP_MATERIAIS['Material']),
                                         )

            index_procurado = SAP_MATERIAIS[SAP_MATERIAIS['Material'] == op_material].index[0]
            op_undmaterial = SAP_MATERIAIS['UM básica'][index_procurado]
            col1.write(SAP_MATERIAIS['Texto breve material'][index_procurado])

            op_qtdmaterial = col1.number_input(label='Quantidade do material',
                                               format='%f',
                                               step=1.,
                                               )
    else:
        op_material = ''
        op_qtdmaterial = ''
        op_undmaterial = ''

    adc_op_mat = col1.button(label='Adicionar Operação (Material)')

    #col2.table(SAP_EQP_filt.head(20))
    col3, col4, col5 = col2.columns([2, 3, 6])
    #   Adicionar operação
    adc_op = col3.button(label='Adicionar Operação')
    #   Limpar operação
    limpar_op = col4.button(label='Limpar Todas as Operações')

    if 'OP_PADRAO' not in st.session_state or limpar_op:
        OP_PADRAO = pd.DataFrame({
            "TEXTO TIPO EQUIPAMENTO": [],
            "TIPO EQUIPAMENTO": [],
            "VARIACAO DESC": [],
            "VARIACAO N4/N5": [],
            "TASK LIST_TRECHO 01": [],
            "TIPO ATIV": [],
            "ROTA?": [],
            "TRECHO 02_TASK LIST": [],
            "ESPECIALIDADE": [],
            "PERIODICIDADE": [],
            "ESTADO MAQ": [],
            "TASK LIST_PARCIAL": [],
            "OPERACAO PADRAO": [],
            "TEXTO DESCRITIVO": [],
            "CT PM": [],
            "DUR_NORMAL_MIN": [],
            "QTD": [],
            "TRABALHO_MIN": [],
            "TASK LIST ORIGINAL": [],
            "TEM MATERIAL?": [],
            "MATERIAL":[],
            "QTD_MATERIAL":[],
            "UND_MATERIAL":[],
            "Planta*": []
        })
        st.session_state.OP_PADRAO = OP_PADRAO
    else:
        OP_PADRAO = st.session_state['OP_PADRAO']

    if adc_op or adc_op_mat:      #  Ao clicar botão de adicionar operação
        for i in range(len(op_eqp_)):

            lista_add_OP_PADRAO = [op_eqp_[i],op_eqp_id[i],np.nan,np.nan,np.nan,op_tipoativ,op_erota,op_trecho2tasklist,op_espec,op_periodicidade,op_estadomaq]
            task_list_parcial_ = (op_tipoativ[0:4]+' '+op_espec[0:3]+' '+op_estadomaq_+' '+op_periodicidade+' '+op_trecho2tasklist).replace('  ', ' ') if op_erota=='NÃO' else (op_tipoativ[0:4]+' '+op_espec[0:3]+' '+op_estadomaq_+' '+op_periodicidade).replace('  ', ' ')
            lista_add_OP_PADRAO.append(task_list_parcial_)
            lista_add_OP_PADRAO = lista_add_OP_PADRAO+[op_oppadrao,op_optextolongo,op_ctpm,op_duracao,op_operad,op_trabtot]

            ## Task list completa (TASK LIST ORIGINAL):
            index_procurado = SAP_EQP_N6[SAP_EQP_N6['ID_SAP_N6'] == op_eqp_id[i]].index
            if index_procurado.empty:
                index_procurado = SAP_EQP_N6[SAP_EQP_N6['ID_SAP_N7'] == op_eqp_id[i]].index
                if index_procurado.empty:
                    index_procurado = SAP_EQP_N6[SAP_EQP_N6['ID_SAP_N8'] == op_eqp_id[i]].index
            loc_inst_i = SAP_EQP_N6['LINHAS / DIAG / SUB PROCESS'][index_procurado[0]]
            loc_inst_desc_i = SAP_EQP_N6['DESC SISTEMAS / ETAPAS PROCESS'][index_procurado[0]]
            task_list_completa_ = task_list_parcial_+' '+str(loc_inst_desc_i)+' '+str(loc_inst_i) if op_erota=='SIM' else task_list_parcial_+' '+str(loc_inst_i)
            ### Check dos 40 caracteres para rotas (pegam local de instalação)
            if len(task_list_completa_)>40:
                task_list_completa_ = task_list_completa_.replace('.', '').replace(' N0', ' ').replace(' DE ', ' ').replace(' DA ', ' ').replace(' DO ', ' ').replace('  ', ' ')
            if len(task_list_completa_) > 40:
                task_list_completa_ = task_list_completa_.replace(' RESFRIAMENTO ', ' RESF ').replace(' RESFRIADOR ', ' RESF ').replace(' REFRIGERACAO ', ' REFRIG ').replace(' EMBALAGEM ', ' EMB ').replace(' EMBALADORA ', ' EMB ').replace(' EMBALADOR ', ' EMB ').replace(' ENCAIXOTAMENTO ', ' ENCX ').replace(' ENCAIXOTADOR ', ' ENCX ').replace(' ENCAIXOTADORA ', ' ENCX ').replace(' EMPILHAMENTO ', ' EMPI ').replace(' APLICACAO ', ' APL ').replace(' PREPARO ', ' PREP ').replace(' PREPARACAO ', ' PREP ').replace(' BISCOITO ', ' BISC ').replace(' WAFER ', ' WAF ').replace(' TRANSPORTE ', ' TRANSP ').replace(' TRANSFERENCIA ', ' TRANSF ').replace(' ARMAZENAGEM ', ' ARMZ ').replace(' ALIMENTACAO ', ' ALIM ').replace(' SEPARACAO ', ' SEP ').replace(' RECEBIMENTO ', ' RECEB ').replace(' MARGARINA ', ' MARG ').replace(' DISTRIBUICAO ', ' DISTRIB ').replace(' GERACAO ', ' GER ').replace(' COMPRIMIDO ', ' COMP ').replace('  ', ' ')
            lista_add_OP_PADRAO = lista_add_OP_PADRAO + [task_list_completa_]

            ## Material
            lista_add_OP_PADRAO = lista_add_OP_PADRAO + [op_temmaterial,op_material, op_qtdmaterial, op_undmaterial]

            ## Salvando dados extras para guardar entradas (com *):
            lista_add_OP_PADRAO = lista_add_OP_PADRAO + [plant_option]

            OP_PADRAO.loc[len(OP_PADRAO)] = lista_add_OP_PADRAO #   Adicionar nova linha
        st.session_state['OP_PADRAO'] = OP_PADRAO

    EQUIPAMENTOS = pd.DataFrame({})  # Inicializar DataFrame vazio
    for i in range(len(OP_PADRAO["TIPO EQUIPAMENTO"])):
        ## Copiar SAP_EQP_N6 apenas na primeira iteração
        if i == 0:
            EQUIPAMENTOS = SAP_EQP_N6[[]].copy()

        ## Filtrar SAP_EQP_N6 com base no ID_SAP_N6, ID_SAP_N7 e ID_SAP_N8
        EQUIPAMENTOS_i = SAP_EQP_N6[
            SAP_EQP_N6['ID_SAP_N6'].isin([OP_PADRAO["TIPO EQUIPAMENTO"][i]])].copy().reset_index(drop=True)
        if not EQUIPAMENTOS_i.empty:
            EQUIPAMENTOS_i['Tipo OP Padrão N6'] = OP_PADRAO["TEXTO TIPO EQUIPAMENTO"][i]

        if EQUIPAMENTOS_i.empty:
            EQUIPAMENTOS_i = SAP_EQP_N6[
                SAP_EQP_N6['ID_SAP_N7'].isin([OP_PADRAO["TIPO EQUIPAMENTO"][i]])].copy().reset_index(drop=True)
            if not EQUIPAMENTOS_i.empty:
                EQUIPAMENTOS_i['Tipo OP Padrão N7'] = OP_PADRAO["TEXTO TIPO EQUIPAMENTO"][i]

        if EQUIPAMENTOS_i.empty:
            EQUIPAMENTOS_i = SAP_EQP_N6[
                SAP_EQP_N6['ID_SAP_N8'].isin([OP_PADRAO["TIPO EQUIPAMENTO"][i]])].copy().reset_index(drop=True)
            if not EQUIPAMENTOS_i.empty:
                EQUIPAMENTOS_i['Tipo OP Padrão N8'] = OP_PADRAO["TEXTO TIPO EQUIPAMENTO"][i]

        EQUIPAMENTOS = pd.concat([EQUIPAMENTOS, EQUIPAMENTOS_i], ignore_index=True, sort=False)
        EQUIPAMENTOS = checagem_df(EQUIPAMENTOS)

    #st.dataframe(EQUIPAMENTOS)#**

    OP_PADRAO.drop_duplicates(subset=None, inplace=True)  # Para remover todas as linhas duplicadas
    OP_PADRAO.sort_values(by=['Planta*','TASK LIST ORIGINAL','TEXTO TIPO EQUIPAMENTO'], inplace=True)      # Ordem com index ordenado com base na task list
    OP_PADRAO.reset_index(drop=True, inplace=True)


    #  NÃO ADICIONAR SE A OPERAÇÃO ESTIVER INCOMPLETA:
    OP_PADRAO.drop(OP_PADRAO[OP_PADRAO['OPERACAO PADRAO'] == ''].index, inplace=True)
    OP_PADRAO.drop(OP_PADRAO[(OP_PADRAO['TRECHO 02_TASK LIST'] == '') & (OP_PADRAO['ROTA?'] == 'NÃO')].index, inplace=True)
    OP_PADRAO.drop(OP_PADRAO[(OP_PADRAO['QTD_MATERIAL'] == 0) & (OP_PADRAO['TEM MATERIAL?'] == 'SIM')].index, inplace=True)
    OP_PADRAO.reset_index(drop=True, inplace=True)


    #   Remover linha de operação específica
    #col2.divider()
    opcoes = ['Selecione o índice da linha'] + [indice for indice in OP_PADRAO.index.tolist()]
    index_remov_edit = col5.selectbox(label='Selecione uma linha para editar ou excluir:', options=opcoes,key='remov_edit')

    def reset():
        st.session_state.remov_edit = 'Selecione o índice da linha'
        st.session_state.PLANT_OPTION = OP_PADRAO["Planta*"][index_remov_edit]
        st.session_state.OP_EQP = OP_PADRAO["TIPO EQUIPAMENTO"][index_remov_edit]
        st.session_state.OP_CTPM = OP_PADRAO["CT PM"][index_remov_edit]
        st.session_state.OP_OPPADRAO = OP_PADRAO["OPERACAO PADRAO"][index_remov_edit]
        st.session_state.OP_OPTEXTOLONGO = OP_PADRAO["TEXTO DESCRITIVO"][index_remov_edit]
        st.session_state.OP_TIPOATIV = OP_PADRAO["TIPO ATIV"][index_remov_edit]
        st.session_state.OP_PERIODICIDADE = OP_PADRAO["PERIODICIDADE"][index_remov_edit]
        st.session_state.OP_ESTADOMAQ = OP_PADRAO["ESTADO MAQ"][index_remov_edit]
        st.session_state.OP_EROTA = OP_PADRAO["ROTA?"][index_remov_edit]
        st.session_state.OP_ESPEC = OP_PADRAO["ESPECIALIDADE"][index_remov_edit]
        st.session_state.OP_DURACAO = OP_PADRAO["DUR_NORMAL_MIN"][index_remov_edit]
        st.session_state.OP_OPERAD = OP_PADRAO["QTD"][index_remov_edit]
        st.session_state.OP_TRECHO2TASKLIST = OP_PADRAO["TRECHO 02_TASK LIST"][index_remov_edit]

        st.session_state.OP_TEMMATERIAL = OP_PADRAO["TEM MATERIAL?"][index_remov_edit]
        st.session_state.OP_MATERIAL = OP_PADRAO["MATERIAL"][index_remov_edit]
        st.session_state.OP_QTDMATERIAL = OP_PADRAO["QTD_MATERIAL"][index_remov_edit]
        #st.session_state.UND_MATERIAL = OP_PADRAO["UND_MATERIAL"][index_remov_edit]

        if index_remov_edit != 'Selecione o índice da linha':
            OP_PADRAO.drop(int(index_remov_edit), inplace=True)
            OP_PADRAO.reset_index(drop=True, inplace=True)


    excluir_linha = col5.button('Selecionar linha', on_click=reset)
    col2.divider()

    st.session_state['OP_PADRAO'] = OP_PADRAO


    # Tabelas de operações
    OP_PADRAO_SHOW = OP_PADRAO[['TASK LIST ORIGINAL','TEXTO TIPO EQUIPAMENTO','TIPO EQUIPAMENTO','OPERACAO PADRAO','TEXTO DESCRITIVO','CT PM','DUR_NORMAL_MIN',
                                'QTD','TRABALHO_MIN','MATERIAL','QTD_MATERIAL','UND_MATERIAL','Planta*']]

    ##   Ajustando planilha op_padrao - sem materiais (OP_PADRAO_OFC)
    OP_PADRAO.sort_values(by=['Planta*','TEXTO TIPO EQUIPAMENTO','ESPECIALIDADE','PERIODICIDADE'],
                          inplace=True)  # Ordem original da op_padrao (por equipamento)

    OP_PADRAO_OFC = OP_PADRAO.copy()
    OP_PADRAO_OFC.drop(columns=['TEM MATERIAL?', 'MATERIAL', 'QTD_MATERIAL', 'UND_MATERIAL'], inplace=True)
    OP_PADRAO_OFC.drop_duplicates(subset=None, inplace=True)  # Para remover todas as linhas duplicadas

    #col2.write(OP_PADRAO_OFC)

    col2.subheader("Listas de Tarefas Geradas", divider='gray', anchor=False)

    col2.markdown("""❗ É recomendável :red[salvar o arquivo sempre que possível] a fim de não haver risco de perder as listas criadas.   
    ❗  Após a conclusão: enviar arquivo para :red[João Ivo] (ter08334@mdb.com.br)
                  """)

    # Download Button
    import io
    buffer1 = io.BytesIO()
    with pd.ExcelWriter(buffer1, engine="xlsxwriter") as writer:
        OP_PADRAO_SHOW.to_excel(writer, sheet_name="Task_Lists", index=False)
        EQUIPAMENTOS.to_excel(writer, sheet_name="EQUIPAMENTOS", index=False)
        OP_PADRAO_OFC.to_excel(writer, sheet_name="op_padrao", index=False)
        # Close the Pandas Excel writer and output the Excel file to the buffer
        writer.close()

        col2.download_button(
            label="Download Listas de Tarefas",
            data=buffer1,
            file_name="Op_Padrao_SAP.xlsx"
        )

    col2.write(OP_PADRAO_SHOW)

    # Tabela de ajustes necessários e observações (revisão)

    OP_PADRAO_SHOW_OBS = pd.DataFrame({'ÍNDICE':[],'OBSERVAÇÃO': []})
    OP_PADRAO_SHOW_AJUS = pd.DataFrame({'ÍNDICE': [], 'AJUSTES NECESSÁRIOS': []})

    for i in range(len(OP_PADRAO_SHOW['TEXTO TIPO EQUIPAMENTO'])):

        ## OBS.:

        ## Sem texto longo  <>
        if pd.isna(OP_PADRAO_SHOW['TEXTO DESCRITIVO'][i]):
            OP_PADRAO_SHOW_OBS.loc[len(OP_PADRAO_SHOW_OBS)] = [i, "VERIFICAR SE TEXTO LONGO (DETALHADO) É NECESSÁRIO"]

        ### TL com máquina parada e peridiocidade menor que 1M  <>
        if ' FUNC ' not in OP_PADRAO_SHOW['TASK LIST ORIGINAL'][i] and OP_PADRAO_SHOW['TASK LIST ORIGINAL'][i].split(' ')[2][-1] == 'D':
            OP_PADRAO_SHOW_OBS.loc[len(OP_PADRAO_SHOW_OBS)] = [i, "MÁQUINA PARADA COM PERIODICIDADE < 1M"]

        ### TL com apenas uma operação  <>
        if len(OP_PADRAO_SHOW[OP_PADRAO_SHOW['TASK LIST ORIGINAL']==OP_PADRAO_SHOW['TASK LIST ORIGINAL'][i]]) <= 1:
            OP_PADRAO_SHOW_OBS.loc[len(OP_PADRAO_SHOW_OBS)] = [i, "TASK LIST COM APENAS UMA OPERAÇÃO"]

        ## AJUSTE:

        ### Suboperação não especifica tipo de atividade? <>
        if isinstance(OP_PADRAO_SHOW['OPERACAO PADRAO'][i], str):
            if not any(name in OP_PADRAO_SHOW['OPERACAO PADRAO'][i] for name in
                       ['INSP', 'TROC', 'REV', 'BLOQ', 'DESBLOQ', 'DOSA', 'SUB', 'VERI', 'DESL', 'LIG', 'CONEC',
                        'DESCON', 'REAL', 'LIMP', 'EFET', 'EXEC', 'TEST', 'ISOL', 'RETI', 'ABR', 'FECH', 'LUB',
                        'ESGOT',
                        'SOLIC', 'ENSA', 'CALI', 'AJUS', 'REAP', 'FIX', 'PARAF', 'INSER', 'COLOC',
                        'AJUS']) and not any(
                letr_2 in OP_PADRAO_SHOW['OPERACAO PADRAO'][i].split()[0][-2:] for letr_2 in ['AR', 'ER', 'IR']):
                OP_PADRAO_SHOW_AJUS.loc[len(OP_PADRAO_SHOW_AJUS)] = [i,"VERIFICAR SE EM 'OPERAÇÃO PADRÃO' É ESPECIFICADO TIPO DE ATIVIDADE DESEMPENHADA (INSP, REVI, ETC)"]

    col2.subheader("Ajustes necessários", divider='gray',anchor=False)
    col2.dataframe(data=OP_PADRAO_SHOW_AJUS,hide_index=True,use_container_width=True)
    col2.subheader("Observações", divider='gray',anchor=False)
    col2.dataframe(data=OP_PADRAO_SHOW_OBS,hide_index=True,use_container_width=True)

    # Pesquisar códigos de equipamentos pelo nome
    st.subheader("🔎 Código de Equipamento por Nome", divider='gray')
    st.markdown("[Criar Lista de Tarefa](#listas-de-tarefas-sap)", unsafe_allow_html=True)
    pesquisar_eqp = st.text_input("Pesquisar Equipamento:")
    pesquisar_eqp = pesquisar_eqp.upper()
    PROC_EQP = SAP_EQP_filt[SAP_EQP_filt['Denominação do objeto técnico'].str.contains(str(pesquisar_eqp))].copy().reset_index(drop=True)
    PROC_EQP = PROC_EQP[['Equipamento','Denominação do objeto técnico','Local de instalação']]
    st.dataframe(data=PROC_EQP, hide_index=True, use_container_width=True)


    # Pesquisar códigos de materiais pelo nome
    st.subheader("🔎 Código de Material por Nome", divider='gray', anchor=False)
    pesquisar_mat = st.text_input("Pesquisar Material:")
    pesquisar_mat = pesquisar_mat.upper()
    PROC_MAT = SAP_MATERIAIS[
        SAP_MATERIAIS['Texto breve material'].str.contains(str(pesquisar_mat))].copy().reset_index(drop=True)
    PROC_MAT = PROC_MAT[['Material', 'Texto breve material']]
    st.dataframe(data=PROC_MAT, hide_index=True, use_container_width=True)

    st.divider()
