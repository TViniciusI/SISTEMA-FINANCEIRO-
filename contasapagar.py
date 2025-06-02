# Desenvolvido por Vinicius Magalhães
import streamlit as st
import pandas as pd
import os
from datetime import datetime, date
from openpyxl import load_workbook

# CONFIGURAÇÃO DE PÁGINA
st.set_page_config(page_title="Sistema Financeiro", page_icon="💰", layout="wide")

# CONSTANTES (os dois arquivos .xlsx devem estar na mesma pasta que este script)
EXCEL_PAGAR = "Contas a pagar 2025 Sistema.xlsx"
EXCEL_RECEBER = "Contas a Receber 2025 Sistema.xlsx"
ANEXOS_DIR = "anexos"

# FUNÇÕES AUXILIARES
def get_sheet_list(excel_path: str):
    """Retorna lista de abas, ignorando aba 'Tutorial' se existir."""
    try:
        wb = pd.ExcelFile(excel_path)
        sheets = [s for s in wb.sheet_names if s.lower() != 'tutorial']
        return sheets
    except Exception:
        return []

def load_data(excel_path: str, sheet_name: str) -> pd.DataFrame:
    """
    Carrega dados da aba especificada, renomeia colunas e calcula status_pagamento.
    Os skiprows=7 posicionam no header correto para as colunas.
    """
    df = pd.read_excel(excel_path, sheet_name=sheet_name, skiprows=7)
    # Se a primeira coluna vier como 'Unnamed...', removemos
    if df.columns[0].lower().startswith('unnamed'):
        df = df.drop(df.columns[0], axis=1)

    # Ajusta nome de colunas conforme a quantidade de colunas detectadas
    if df.shape[1] == 10:
        df.columns = [
            'data_nf', 'forma_pagamento', 'fornecedor', 'os',
            'vencimento', 'valor', 'estado', 'situacao', 'boleto', 'comprovante'
        ]
    elif df.shape[1] == 8:
        df.columns = [
            'data_nf', 'forma_pagamento', 'fornecedor', 'os',
            'vencimento', 'valor', 'estado', 'situacao'
        ]

    # Remove linhas sem fornecedor ou valor
    df = df.dropna(subset=['fornecedor', 'valor']).reset_index(drop=True)

    # Conversões de tipos
    df['vencimento'] = pd.to_datetime(df['vencimento'], errors='coerce')
    df['valor'] = pd.to_numeric(df['valor'], errors='coerce')

    # Cálculo de status_pagamento
    status_list = []
    hoje = datetime.now().date()
    for _, row in df.iterrows():
        pago = False
        if sheet_name.lower().startswith('contas a pagar'):
            if row['estado'] == 'Pago':
                pago = True
        else:
            if row['estado'] == 'Recebido':
                pago = True

        data_venc = row['vencimento'].date() if not pd.isna(row['vencimento']) else None
        if pago:
            status_list.append('Em Dia')
        else:
            if data_venc:
                if data_venc < hoje:
                    status_list.append('Em Atraso')
                else:
                    status_list.append('A Vencer')
            else:
                status_list.append('Sem Data')

    df['status_pagamento'] = status_list
    return df

def save_data(excel_path: str, sheet_name: str, df: pd.DataFrame):
    """
    Salva as colunas valor, estado, situacao e vencimento de volta
    na planilha Excel, respeitando a posição original (linha +8).
    """
    wb = load_workbook(excel_path)
    ws = wb[sheet_name]
    for i, row in df.iterrows():
        ws.cell(row=i + 8, column=6, value=row['valor'])
        ws.cell(row=i + 8, column=7, value=row['estado'])
        ws.cell(row=i + 8, column=8, value=row['situacao'])
        ws.cell(row=i + 8, column=5, value=row['vencimento'])
    wb.save(excel_path)

def add_record(excel_path: str, sheet_name: str, record: dict):
    """
    Adiciona um novo registro na próxima linha disponível da aba especificada.
    Campos extras (boleto e comprovante) também serão gravados se existirem.
    """
    wb = load_workbook(excel_path)
    ws = wb[sheet_name]
    next_row = ws.max_row + 1

    valores = [
        record.get('data_nf', ''),
        record.get('forma_pagamento', ''),
        record.get('fornecedor', ''),
        record.get('os', ''),
        record.get('vencimento', ''),
        record.get('valor', ''),
        record.get('estado', ''),
        record.get('situacao', ''),
        record.get('boleto', ''),
        record.get('comprovante', '')
    ]

    for col_idx, val in enumerate(valores, start=1):
        ws.cell(row=next_row, column=col_idx, value=val)

    wb.save(excel_path)

# Garante que as pastas de anexos existam
for pasta in ['Contas a Pagar', 'Contas a Receber']:
    os.makedirs(os.path.join(ANEXOS_DIR, pasta), exist_ok=True)

# SIDEBAR
st.sidebar.markdown("# 📂 Navegação")
page = st.sidebar.radio("", ['Dashboard', 'Contas a Pagar', 'Contas a Receber'])

# HEADER PRINCIPAL
st.markdown(
    "<h1 style='text-align: center; color: #4B8BBE; font-size:32px;'>💼 Sistema Financeiro 2025</h1>",
    unsafe_allow_html=True
)
st.markdown(
    "<p style='text-align: center; color: #555; font-size:16px;'>"
    "Gestão eficiente de contas a pagar e receber com indicadores e gráficos intuitivos."
    "</p>",
    unsafe_allow_html=True
)
st.markdown("---")

# =========================
#    SEÇÃO: DASHBOARD
# =========================
if page == 'Dashboard':
    st.subheader("📊 Painel de Controle Financeiro")

    # Obtém as abas de cada arquivo
    sheets_p = get_sheet_list(EXCEL_PAGAR)
    sheets_r = get_sheet_list(EXCEL_RECEBER)

    # -------------------
    # KPIs Contas a Pagar
    # -------------------
    if sheets_p:
        df_all_p = pd.concat([load_data(EXCEL_PAGAR, s) for s in sheets_p], ignore_index=True)
        total_p = df_all_p['valor'].sum()
        pago_p = df_all_p[df_all_p['estado'] == 'Pago']['valor'].sum()
        aberto_p = df_all_p[df_all_p['estado'] == 'Em Aberto']['valor'].sum()
        vencido_p = df_all_p[
            (df_all_p['estado'] == 'Em Aberto') &
            (df_all_p['vencimento'] < datetime.now())
        ]['valor'].sum()

        st.markdown(
            "<div style='background-color: #E8F8F5; padding:10px; border-radius:5px;'>"
            "<strong>Contas a Pagar - KPIs</strong></div>",
            unsafe_allow_html=True
        )
        kp1, kp2, kp3, kp4 = st.columns(4)
        kp1.metric("Total a Pagar", f"R$ {total_p:,.2f}")
        kp2.metric("Pago", f"R$ {pago_p:,.2f}")
        kp3.metric("Em Aberto", f"R$ {aberto_p:,.2f}")
        kp4.metric("Vencido", f"R$ {vencido_p:,.2f}")
    else:
        st.warning("Nenhuma aba encontrada para Contas a Pagar.")
    st.markdown("---")

    # ----------------------
    # KPIs Contas a Receber
    # ----------------------
    if sheets_r:
        df_all_r = pd.concat([load_data(EXCEL_RECEBER, s) for s in sheets_r], ignore_index=True)
        total_r = df_all_r['valor'].sum()
        rec_r = df_all_r[df_all_r['estado'] == 'Recebido']['valor'].sum()
        arec_r = df_all_r[df_all_r['estado'] == 'A Receber']['valor'].sum()
        atras_r = df_all_r[
            (df_all_r['estado'] == 'A Receber') &
            (df_all_r['vencimento'] < datetime.now())
        ]['valor'].sum()

        st.markdown(
            "<div style='background-color: #FEF9E7; padding:10px; border-radius:5px;'>"
            "<strong>Contas a Receber - KPIs</strong></div>",
            unsafe_allow_html=True
        )
        kr1, kr2, kr3, kr4 = st.columns(4)
        kr1.metric("Total a Receber", f"R$ {total_r:,.2f}")
        kr2.metric("Recebido", f"R$ {rec_r:,.2f}")
        kr3.metric("A Receber", f"R$ {arec_r:,.2f}")
        kr4.metric("Atrasado", f"R$ {atras_r:,.2f}")
    else:
        st.warning("Nenhuma aba encontrada para Contas a Receber.")
    st.markdown("---")

    # -------------------------------
    # Gráficos Mensais (Pagar vs Receber)
    # -------------------------------
    col1, col2 = st.columns(2)

    if sheets_p:
        monthly_p = {s: load_data(EXCEL_PAGAR, s)['valor'].sum() for s in sheets_p}
        with col1:
            st.markdown("<div style='text-align:center;'><strong>Gastos Mensais</strong></div>", unsafe_allow_html=True)
            st.bar_chart(pd.Series(monthly_p), use_container_width=True)

    if sheets_r:
        monthly_r = {s: load_data(EXCEL_RECEBER, s)['valor'].sum() for s in sheets_r}
        with col2:
            st.markdown("<div style='text-align:center;'><strong>Receitas Mensais</strong></div>", unsafe_allow_html=True)
            st.bar_chart(pd.Series(monthly_r), use_container_width=True)

    # -----------------------------------------------------
    # NOVA SEÇÃO: Exportar Planilhas (Download dos Arquivos)
    # -----------------------------------------------------
    st.markdown("---")
    st.subheader("💾 Exportar Planilhas Originais")
    export_col1, export_col2 = st.columns(2)

    with export_col1:
        try:
            with open(EXCEL_PAGAR, "rb") as f:
                bytes_data_p = f.read()
            st.download_button(
                label="Exportar Contas a Pagar",
                data=bytes_data_p,
                file_name=EXCEL_PAGAR,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except FileNotFoundError:
            st.error(f"Arquivo '{EXCEL_PAGAR}' não encontrado.")

    with export_col2:
        try:
            with open(EXCEL_RECEBER, "rb") as f:
                bytes_data_r = f.read()
            st.download_button(
                label="Exportar Contas a Receber",
                data=bytes_data_r,
                file_name=EXCEL_RECEBER,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except FileNotFoundError:
            st.error(f"Arquivo '{EXCEL_RECEBER}' não encontrado.")

# ===========================
#    SEÇÃO: CONTAS A PAGAR
#    ou CONTAS A RECEBER
# ===========================
else:
    st.subheader(f"🗂️ {page}")
    excel_path = EXCEL_PAGAR if page == 'Contas a Pagar' else EXCEL_RECEBER
    sheets = get_sheet_list(excel_path)

    if not sheets:
        st.error(f"Arquivo '{excel_path}' não encontrado ou sem abas válidas.")
        st.stop()

    # Seleção de Mês (aba)
    aba = st.selectbox("Selecione o mês:", sheets)
    df = load_data(excel_path, aba)

    if df.empty:
        st.info("Nenhum registro encontrado para este mês.")
    else:
        # ====================
        #      FILTROS
        # ====================
        colf1, colf2 = st.columns(2)
        with colf1:
            # Remove NaNs e converte para string antes de ordenar
            fornecedores_list = df['fornecedor'].dropna().astype(str).unique().tolist()
            forn = st.selectbox("Fornecedor", ['Todos'] + sorted(fornecedores_list))
        with colf2:
            # Remove NaNs e converte para string antes de ordenar
            estados_list = df['estado'].dropna().astype(str).unique().tolist()
            status_sel = st.selectbox("Status/Estado", ['Todos'] + sorted(estados_list))

        # Aplica filtros apenas se não forem "Todos"
        if forn != 'Todos':
            df = df[df['fornecedor'] == forn]
        if status_sel != 'Todos':
            df = df[df['estado'] == status_sel]

        st.markdown("<hr style='border:1px solid #ddd;'>", unsafe_allow_html=True)

        # Se após filtros o DataFrame ficar vazio, exibimos aviso e encerramos essa seção
        if df.empty:
            st.warning("Nenhum registro corresponde aos filtros selecionados.")
        else:
            # ====================
            #   LISTA DE LANÇAMENTOS
            # ====================
            st.subheader("📋 Lista de Lançamentos")
            st.table(df[['data_nf', 'fornecedor', 'valor', 'vencimento', 'status_pagamento']])
            st.markdown("---")

            # ====================
            #    EDITAR REGISTRO
            # ====================
            st.subheader("✏️ Editar Registro")
            idx = st.number_input(
                "Índice da linha:", min_value=0, max_value=len(df) - 1, step=1
            )
            rec = df.iloc[idx]

            colv1, colv2 = st.columns(2)
            with colv1:
                new_valor = st.number_input(
                    "Valor:", value=float(rec['valor']), key="valores"
                )
                default_date = rec['vencimento'].date() if pd.notna(rec['vencimento']) else date.today()
                new_venc = st.date_input("Vencimento:", value=default_date, key="vencimento")
            with colv2:
                # CORREÇÃO: cria lista única de 'estado' e 'situacao' e garante índice válido
                estado_list = df['estado'].dropna().astype(str).unique().tolist()
                try:
                    estado_index = estado_list.index(str(rec['estado']))
                except ValueError:
                    estado_index = 0
                new_estado = st.selectbox(
                    "Estado:",
                    options=estado_list,
                    index=estado_index,
                    key="estado"
                )

                situ_list = df['situacao'].dropna().astype(str).unique().tolist()
                try:
                    situ_index = situ_list.index(str(rec['situacao']))
                except ValueError:
                    situ_index = 0
                new_sit = st.selectbox(
                    "Situação:",
                    options=situ_list,
                    index=situ_index,
                    key="situacao"
                )

            if st.button("💾 Salvar Alterações"):
                df.loc[df.index[idx], ['valor', 'vencimento', 'estado', 'situacao']] = [
                    new_valor, pd.to_datetime(new_venc), new_estado, new_sit
                ]
                save_data(excel_path, aba, df)
                st.success("Registro atualizado no Excel.")

            st.markdown("---")

            # ====================
            #    ANEXAR DOCUMENTOS
            # ====================
            st.subheader("📎 Anexar Documentos")
            uploaded = st.file_uploader(
                "Selecione o arquivo (pdf/jpg/png):", type=['pdf', 'jpg', 'png'], key=f"up_{page}_{aba}_{idx}"
            )
            if uploaded:
                destino = os.path.join(ANEXOS_DIR, page, f"{page}_{aba}_{idx}_{uploaded.name}")
                with open(destino, 'wb') as f:
                    f.write(uploaded.getbuffer())
                st.success(f"Documento salvo em: {destino}")

            st.markdown("---")

            # ====================
            #   ADICIONAR NOVA CONTA
            # ====================
            st.subheader("➕ Adicionar Nova Conta")
            with st.form(key="form_add"):
                coln1, coln2 = st.columns(2)
                with coln1:
                    data_nf = st.date_input("Data N/F:", value=date.today())
                    forma_pag = st.text_input("Forma de Pagamento:")
                    forn_new = st.text_input("Fornecedor:")
                with coln2:
                    os_new = st.text_input("OS Interna:")
                    venc_new = st.date_input("Data de Vencimento:", value=date.today())
                    valor_new = st.number_input("Valor (R$):", min_value=0.0, format="%.2f")

                if page == 'Contas a Pagar':
                    estado_opt = ['Em Aberto', 'Pago']
                    situ_opt = ['Em Atraso', 'Pago', 'Em Aberto']
                else:
                    estado_opt = ['A Receber', 'Recebido']
                    situ_opt = ['Em Atraso', 'Recebido', 'A Receber']

                estado_new = st.selectbox("Estado:", options=estado_opt)
                situ_new = st.selectbox("Situação:", options=situ_opt)
                boleto_file = st.file_uploader("Boleto (opcional):", type=['pdf', 'jpg', 'png'])
                comprov_file = st.file_uploader("Comprovante (opcional):", type=['pdf', 'jpg', 'png'])
                submit_add = st.form_submit_button("➕ Adicionar Conta")

            if submit_add:
                boleto_path = ''
                comprov_path = ''
                if boleto_file:
                    boleto_path = os.path.join(ANEXOS_DIR, page, f"{page}_{aba}_boleto_{boleto_file.name}")
                    with open(boleto_path, 'wb') as f:
                        f.write(boleto_file.getbuffer())
                if comprov_file:
                    comprov_path = os.path.join(ANEXOS_DIR, page, f"{page}_{aba}_comprov_{comprov_file.name}")
                    with open(comprov_path, 'wb') as f:
                        f.write(comprov_file.getbuffer())

                record = {
                    'data_nf': data_nf,
                    'forma_pagamento': forma_pag,
                    'fornecedor': forn_new,
                    'os': os_new,
                    'vencimento': venc_new,
                    'valor': valor_new,
                    'estado': estado_new,
                    'situacao': situ_new,
                    'boleto': boleto_path,
                    'comprovante': comprov_path
                }
                add_record(excel_path, aba, record)
                st.success("Nova conta adicionada com sucesso!")

            # ------------------------------------
            # SEÇÃO OPCIONAL: Exportar Aba Atual
            # ------------------------------------
            st.markdown("---")
            st.subheader("💾 Exportar Aba Atual")
            try:
                # Antes de exportar, certifique-se de salvar as alterações feitas na aba atual
                save_data(excel_path, aba, df)
                with open(excel_path, "rb") as f:
                    excel_bytes = f.read()
                st.download_button(
                    label=f"Exportar '{aba}' para Excel",
                    data=excel_bytes,
                    file_name=f"{page} - {aba}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Falha ao preparar o download: {e}")

# RODAPÉ
st.markdown(
    "<p style='text-align:center; font-size:12px; color:gray;'>"
    "Desenvolvido por Vinicius Magalhães</p>",
    unsafe_allow_html=True
)
