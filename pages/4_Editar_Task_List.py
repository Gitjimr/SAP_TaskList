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

    st.header("Editar Lista de Tarefas SAP", anchor=False)

    lista_tl = list(SAP_TL['Grupo'].unique())
    lista_tl.sort()

    def delete_tl_edit_from_session_state():
        if 'tl_edit' in st.session_state:
            del st.session_state['tl_edit']
        if 'cab_edit' in st.session_state:
            del st.session_state['cab_edit']


    tl_selec = st.selectbox(label='Selecione o código da lista',
                            options=lista_tl,
                            on_change=delete_tl_edit_from_session_state)

    TL_EDIT = SAP_TL[SAP_TL['Grupo']==str(tl_selec)].copy()

    TL_EDIT['Operação'] = pd.to_numeric(TL_EDIT['Operação'], downcast='integer')
    TL_EDIT['Suboperação'] = pd.to_numeric(TL_EDIT['Suboperação'], downcast='integer')

    TL_EDIT.reset_index(drop=True, inplace=True)
    

    SAP_CTPM_filt = SAP_CTPM[SAP_CTPM['Cen.'] == str(TL_EDIT['Centro planejamento'][0])]
    lista_ctpms = SAP_CTPM_filt['CenTrab'].unique().tolist()
    SAP_EQP_filt = SAP_EQP[SAP_EQP['Local de instalação'].str.contains(str(TL_EDIT['Local instalação operação'].iloc[0]))]
    lista_eqps = list(SAP_EQP_filt['Equipamento'].unique())
    if not lista_eqps:      #   Caso a lista venha vazia, pegar operação da segunda linha (para não lub a primeira é o cabeçalho)
        try:
            SAP_EQP_filt = SAP_EQP[
                SAP_EQP['Local de instalação'].str.contains(str(TL_EDIT['Local instalação operação'].iloc[1]))]
            lista_eqps = list(SAP_EQP_filt['Equipamento'].unique())
        except:
            pass


    if 'cab_edit' not in st.session_state:
        CAB_EDIT = {'Grupo': TL_EDIT['Grupo'][0], 'Descrição': TL_EDIT['Descrição'][0],
                    'Centro de trabalho': TL_EDIT['Centro de trabalho'][0]}
    else:
        CAB_EDIT = st.session_state['cab_edit']

    if 'tl_edit' in st.session_state:
        TL_EDIT = st.session_state['tl_edit']

    def create_TL_EDIT_editor(TL_EDIT):

        TL_EDIT_ = st.data_editor(
            TL_EDIT,
            column_order=[
                'Grupo', 'Descrição', 'Operação', 'Suboperação',
                'Txt.breve operação',
                'Centro de trabalho', 'Duração normal', 'Equipamento operação'
            ],
            disabled=["Grupo", "Operação", "Suboperação", "Descrição"],
            column_config={
                "Duração normal": st.column_config.NumberColumn(
                    min_value=1,
                    required =True,
                    format="%d",
                ),
                "Txt.breve operação": st.column_config.TextColumn(
                    required=True,
                    max_chars=40
                ),
                "Equipamento operação": st.column_config.TextColumn(
                    required=False if TL_EDIT["Equipamento operação"].isna().all() else True,
                    max_chars=10
                ),
                "Centro de trabalho": st.column_config.TextColumn(
                    required=True,
                    max_chars=8
                )
            }
        )
        return TL_EDIT_

    def add_new_row_to_TL_EDIT(TL_EDIT):

        # Adiciona uma nova linha vazia ao DataFrame
        new_index = TL_EDIT.index[-1] + 1 if not TL_EDIT.empty else 0
        TL_EDIT.loc[new_index] = [None] * TL_EDIT.shape[1]

        # Copia valores da penúltima linha para a nova linha
        TL_EDIT.loc[new_index, 'Duração normal'] = '0'
        TL_EDIT.loc[new_index, 'Grupo'] = TL_EDIT['Grupo'].iloc[-2]
        TL_EDIT.loc[new_index, 'Descrição'] = TL_EDIT['Descrição'].iloc[-2]
        TL_EDIT.loc[new_index, 'Centro planejamento'] = TL_EDIT['Centro planejamento'].iloc[-2]
        TL_EDIT.loc[new_index, 'Local instalação operação'] = TL_EDIT['Local instalação operação'].iloc[-2]
        TL_EDIT.loc[new_index, 'ID externo'] = TL_EDIT['ID externo'].iloc[-2]
        TL_EDIT.loc[new_index, 'Grupo planej.'] = TL_EDIT['Grupo planej.'].iloc[-2]
        TL_EDIT.loc[new_index, 'Tipo de roteiro'] = TL_EDIT['Tipo de roteiro'].iloc[-2]
        TL_EDIT.loc[new_index, 'Numerador de grupos'] = TL_EDIT['Numerador de grupos'].iloc[-2]

        if TL_EDIT["Suboperação"].isna().all():
            TL_EDIT.loc[new_index, 'Operação'] = TL_EDIT['Operação'].iloc[-2] + 10
        else:
            TL_EDIT.loc[new_index, 'Operação'] = TL_EDIT['Operação'].iloc[-2]
            TL_EDIT.loc[new_index, 'Suboperação'] = TL_EDIT['Suboperação'].iloc[-2] + 10

        return TL_EDIT


    def delete_row_from_TL_EDIT(TL_EDIT, index_to_delete):

        if index_to_delete in TL_EDIT.index:
            TL_EDIT = TL_EDIT.drop(index_to_delete).reset_index(drop=True)

        if TL_EDIT["Suboperação"].isna().all():
            value = 10
            for i in range(0, len(TL_EDIT)):
                TL_EDIT.at[i, 'Operação'] = value
                value += 10
        else:
            value = 10
            for i in range(1, len(TL_EDIT)):
                TL_EDIT.at[i, 'Suboperação'] = value
                value += 10

        return TL_EDIT


    def editar_dados_cab(TL_EDIT, CAB_EDIT):

        TL_EDIT["Descrição"] = CAB_EDIT["Descrição"]
        TL_EDIT["Centro de trabalho"] = CAB_EDIT["Centro de trabalho"]

        if not TL_EDIT["Suboperação"].isna().all() and (pd.isna(TL_EDIT["Suboperação"][0]) or TL_EDIT["Suboperação"][0] is None):
            TL_EDIT["Txt.breve operação"][0] = CAB_EDIT["Descrição"]

        return TL_EDIT


    @st.experimental_dialog("📝 Editar Cabeçalho da Lista de Tarefa")
    def edit_header_tl(CAB_EDIT,lista_ctpms):
        try:
            index_procurado = lista_ctpms.index(CAB_EDIT['Centro de trabalho'])
        except:
            index_procurado = 0

        cab_edit_desc = st.text_input(label='Enter some text', value=CAB_EDIT["Descrição"])
        cab_edit_ctpm = st.selectbox(
        label='Centro de Trabalho',
        options=lista_ctpms,
        index = index_procurado)
        submit_button = st.button(label='Submit')
        if submit_button:
            st.session_state.cab_edit = {"Grupo":CAB_EDIT["Grupo"], "Descrição": cab_edit_desc, "Centro de trabalho": cab_edit_ctpm}
            return 1
        else:
            return 0


    if tl_selec is not None:  # Caso o df esteja vazio

        st.subheader('Editar cabeçalho', divider='gray', anchor=False)

        try:
            index_procurado = lista_ctpms.index(CAB_EDIT['Centro de trabalho'])
        except:
            index_procurado = 0
        cab_edit_desc = st.text_input(label='Enter some text', value=CAB_EDIT["Descrição"])
        cab_edit_ctpm = st.selectbox(
            label='Centro de Trabalho',
            options=lista_ctpms,
            index=index_procurado)
        submit_button = st.button(label='Editar',use_container_width=True)
        if submit_button:
            st.session_state.cab_edit = {"Grupo": CAB_EDIT["Grupo"], "Descrição": cab_edit_desc,
                                         "Centro de trabalho": cab_edit_ctpm}
            if 'cab_edit' not in st.session_state:
                st.session_state.cab_edit = CAB_EDIT
            else:
                CAB_EDIT = st.session_state['cab_edit']
            st.session_state["button_pressed"] = True#********
            TL_EDIT = editar_dados_cab(TL_EDIT,CAB_EDIT)#********


        st.subheader('Editar operações', divider='gray', anchor=False)

        col1, col2, col3 = st.columns([2, 4, 2])

        if col1.button('Criar linha', use_container_width=True):
            if 'tl_edit' not in st.session_state:
                st.session_state.tl_edit = TL_EDIT
            else:
                TL_EDIT = st.session_state['tl_edit']
            st.session_state["button_pressed"] = True
            TL_EDIT = add_new_row_to_TL_EDIT(TL_EDIT)

        indices_options = [i for i in range(len(TL_EDIT)) if i != 0] if not TL_EDIT["Suboperação"].isna().all() else [i for i in range(len(TL_EDIT))]
        index_to_delete = col2.selectbox(label='Selecione uma linha para excluir:',
                                      placeholder="Selecione o índice da linha",
                                    options=indices_options
                                         )

        if col3.button('Deletar linha', use_container_width=True):
            if 'tl_edit' not in st.session_state:
                st.session_state.tl_edit = TL_EDIT
            else:
                TL_EDIT = st.session_state['tl_edit']
            st.session_state["button_pressed"] = True
            TL_EDIT = delete_row_from_TL_EDIT(TL_EDIT, index_to_delete)


        tb_rev = {'ÍNDICE': [], 'REVISÃO': []}

        if st.session_state.get("button_pressed", False):

            TL_EDIT = create_TL_EDIT_editor(TL_EDIT)
            #st.write(CAB_EDIT)
            #st.write(TL_EDIT)#**************

            st.session_state.tl_edit = TL_EDIT
            #st.session_state.cab_edit = CAB_EDIT

            #   Revisão da lista editada:

            tb_rev = {'ÍNDICE': [], 'REVISÃO': []}

            for i in range(len(TL_EDIT['Grupo'])):
                try:
                    if TL_EDIT['Centro de trabalho'][i] not in lista_ctpms:
                        tb_rev['ÍNDICE'].append(i)
                        tb_rev['REVISÃO'].append('CENTRO DE TRABALHO NÃO EXISTE NA UNIDADE SELECIONADA.')
                except:
                    tb_rev['ÍNDICE'].append(i)
                    tb_rev['REVISÃO'].append('CENTRO DE TRABALHO NÃO EXISTE NA UNIDADE SELECIONADA.')

                try:
                    if len(str(TL_EDIT['Txt.breve operação'][i])) > 40:
                        tb_rev['ÍNDICE'].append(i)
                        tb_rev['REVISÃO'].append(
                            'Txt.breve operação: QUANTIDADE DE CARACTERES ULTRAPASSANDO O MÁXIMO.')

                    if TL_EDIT['Txt.breve operação'][i] is None:
                        tb_rev['ÍNDICE'].append(i)
                        tb_rev['REVISÃO'].append(
                            'Txt.breve operação: OPERAÇÃO VAZIA.')
                except:
                    tb_rev['ÍNDICE'].append(i)
                    tb_rev['REVISÃO'].append('ERRO NA COLUNA Txt.breve operação')

                try:
                    if TL_EDIT['Duração normal'][i] is None or TL_EDIT['Duração normal'][i]==0 or TL_EDIT['Duração normal'][i]=='0' or TL_EDIT['Duração normal'][i]=='':
                        tb_rev['ÍNDICE'].append(i)
                        tb_rev['REVISÃO'].append(
                            'DURAÇÃO NÃO INSERIDA.')
                except:
                    tb_rev['ÍNDICE'].append(i)
                    tb_rev['REVISÃO'].append('ERRO NA COLUNA DE DURAÇÃO')

                try:
                    if TL_EDIT['Equipamento operação'][i] is None and not TL_EDIT["Equipamento operação"].isna().all() and i != 0:
                        tb_rev['ÍNDICE'].append(i)
                        tb_rev['REVISÃO'].append(
                            'ROTA: NECESSÁRIO INCLUIR CÓDIGO SAP DO EQUIPAMENTO.')
                    if str(TL_EDIT['Equipamento operação'][i]) not in lista_eqps and not TL_EDIT["Equipamento operação"].isna().all() and i>0:
                        tb_rev['ÍNDICE'].append(i)
                        tb_rev['REVISÃO'].append(
                            'ROTA: CÓDIGO DO EQUIPAMENTO NÃO IDENTIFICADO PARA O LOCAL DE INSTALAÇAO DA LISTA.')
                except:
                    tb_rev['ÍNDICE'].append(i)
                    tb_rev['REVISÃO'].append('ERRO NA COLUNA DE EQUIPAMENTO OPERAÇÃO')

                if isinstance(TL_EDIT['Txt.breve operação'][i], str):
                    if not any(name in TL_EDIT['Txt.breve operação'][i] for name in
                               ['INSP', 'TROC', 'REV', 'BLOQ', 'DESBLOQ', 'DOSA', 'SUB', 'VERI', 'DESL', 'LIG', 'CONEC',
                                'DESCON', 'REAL', 'LIMP', 'EFET', 'EXEC', 'TEST', 'ISOL', 'RETI', 'ABR', 'FECH', 'LUB',
                                'ESGOT',
                                'SOLIC', 'ENSA', 'CALI', 'AJUS', 'REAP', 'FIX', 'PARAF', 'INSER', 'COLOC',
                                'AJUS']) and not any(
                        letr_2 in TL_EDIT['Txt.breve operação'][i].split()[0][-2:] for letr_2 in ['AR', 'ER', 'IR']):
                        tb_rev['ÍNDICE'].append(i)
                        tb_rev['REVISÃO'].append(
                            'OPERACAO PADRAO: VERIFICAR ORTOGRAFIA OU SE HÁ A AÇÃO DESEMPENHADA (INSP, REVI, ETC). CASO ERRO PERSISTA, COLOQUE UM VERBO NO INÍCIO.')


            col4, col5 = st.columns([4, 2])

            col4.subheader('Revisões necessárias por linha da tabela', divider='gray', anchor=False)

            df_tb_rev = pd.DataFrame(tb_rev)

            col4.dataframe(data=df_tb_rev, hide_index=True, use_container_width=True)

            # Gerar template de carga do SAP:

            TL_EDIT_TEMPLATE = {'Chave para grupo de listas de tarefas': TL_EDIT['Grupo'], 'Data fixada': ['15.05.2024']*len(TL_EDIT['Grupo']),
                                'Centro': TL_EDIT['Centro planejamento'], 'Numerador de grupos': [np.nan]*len(TL_EDIT['Grupo']), 'Entrada': [i for i in range(len(TL_EDIT['Grupo']))], 'Nº operação': TL_EDIT['Operação'],
                                'Suboperação': TL_EDIT['Suboperação'], 'Txt.breve operação': TL_EDIT['Txt.breve operação'].str.upper(), 'Trabalho da operação': TL_EDIT['Duração normal'],
                                'Unidade de trabalho': ['MIN']*len(TL_EDIT['Grupo']), 'Núm.capacidade necessária': [1]*len(TL_EDIT['Grupo']),
                                'Duração normal da operação': TL_EDIT['Duração normal'], 'Unidade da duração normal': ['MIN']*len(TL_EDIT['Grupo']),
                                'Chave de cálculo': [2]*len(TL_EDIT['Grupo']), 'Porcentagem de trabalho': [100]*len(TL_EDIT['Grupo']), 'Fator de execução': [1]*len(TL_EDIT['Grupo']),
                                'Nº equipamento': TL_EDIT['Equipamento operação'], 'A&D: ID externo da lista de tarefas': TL_EDIT['ID externo']}

            TL_EDIT_TEMPLATE = pd.DataFrame(TL_EDIT_TEMPLATE)

            ##  Somando tempo para não LUB
            if not TL_EDIT_TEMPLATE["Suboperação"].isna().all():
                soma_tempo = TL_EDIT_TEMPLATE['Duração normal da operação'].astype(int).iloc[1:].sum()
                TL_EDIT_TEMPLATE.loc[0, 'Duração normal da operação'] = str(soma_tempo)
                TL_EDIT_TEMPLATE.loc[0, 'Trabalho da operação'] = str(soma_tempo)

            # Download Button
            col5.subheader("⬇️ Download", divider='gray', anchor=False)

            col5.markdown("""❗  Após a conclusão: enviar arquivo para :red[João Ivo] (ter08334@mdb.com.br)
                                      """)

            #st.write(TL_EDIT_TEMPLATE)

            import io

            plant_option = TL_EDIT['Centro planejamento'][0]
            buffer5 = io.BytesIO()

            with pd.ExcelWriter(buffer5, engine="xlsxwriter") as writer:

                TL_EDIT_TEMPLATE.to_excel(writer, sheet_name="TL_EDIT_TEMPLATE", index=False)
                try:
                    df_tb_rev.to_excel(writer, sheet_name="REVISÕES", index=False)
                except:
                    pass
                try:
                    CAB_EDIT_save = pd.DataFrame(CAB_EDIT,index=[0])
                    CAB_EDIT_save.to_excel(writer, sheet_name="CAB_EDIT_TEMPLATE", index=False)
                except:
                    pass
                # Close the Pandas Excel writer and output the Excel file to the buffer
                writer.close()

                col5.download_button(
                    label="Download Lista de Tarefas Revisada",
                    data=buffer5,
                    file_name="EDIT_TL_OP_"+ ".xlsx",
                    use_container_width=True,
                    # disabled=not df_tb_rev.empty

                )






