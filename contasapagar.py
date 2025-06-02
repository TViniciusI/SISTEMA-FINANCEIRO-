# Desenvolvido por Vinicius Magalhães
import streamlit as st
import pandas as pd
import os
from datetime import datetime, date
from openpyxl import load_workbook

# CONFIGURAÇÃO DE PÁGINA
st.set_page_config(
    page_title="💼 Sistema Financeiro 2025",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ====================================================================
#  Autenticação simples (sem bibliotecas externas) com layout personalizado
# ====================================================================

VALID_USERS = {
    "Vinicius": "vinicius4223",
    "Flavio": "1234",
}

def check_login(username: str, password: str) -> bool:
    return VALID_USERS.get(username) == password

# Inicializa estado de sessão
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = ""

# Se não estiver logado, exibe formulário de login estilizado
if not st.session_state.logged_in:
    # Injetar CSS para estilizar o card de login
    st.markdown(
        """
        <style>
        .login-container {
            display: flex;
            justify-content: center;
            align-items: center;
            height: 60vh;
        }
        .login-card {
            background-color: #ffffff;
            padding: 2rem;
            border-radius: 12px;
            box-shadow: 0 4px 16px rgba(0, 0, 0, 0.1);
            max-width: 380px;
            width: 100%;
        }
        .login-card h2 {
            text-align: center;
            color: #4B8BBE;
            margin-bottom: 1.5rem;
            font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        }
        .login-card .stTextInput>div>div>input {
            padding: 0.5rem 0.75rem;
            border: 1px solid #ccc;
            border-radius: 6px;
            width: 100%;
            margin-bottom: 1rem;
            font-size: 1rem;
        }
        .login-card .stButton>button {
            width: 100%;
            padding: 0.6rem;
            background-color: #4B8BBE;
            color: #ffffff;
            font-size: 1rem;
            border: none;
            border-radius: 6px;
            cursor: pointer;
        }
        .login-card .stButton>button:hover {
            background-color: #3A6F9E;
        }
        .login-error {
            color: #D90429;
            font-weight: bold;
            text-align: center;
            margin-top: 0.5rem;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    # Container centralizado
    st.markdown('<div class="login-container">', unsafe_allow_html=True)
    st.markdown('<div class="login-card">', unsafe_allow_html=True)

    st.markdown("<h2>🔒 Acesso Restrito</h2>", unsafe_allow_html=True)

    # Formulário de login
    with st.form("login_form", clear_on_submit=False):
        username_input = st.text_input("Usuário:")
        password_input = st.text_input("Senha:", type="password")
        login_button = st.form_submit_button("Entrar")

        if login_button:
            if check_login(username_input, password_input):
                st.session_state.logged_in = True
                st.session_state.username = username_input
                # Após marcar logged_in, o Streamlit recarrega a página automaticamente
            else:
                st.markdown('<div class="login-error">Usuário ou senha inválidos.</div>', unsafe_allow_html=True)

    st.markdown('</div></div>', unsafe_allow_html=True)
    st.stop()

# Usuário já está autenticado
logged_user = st.session_state.username

# Botão de logout
def logout():
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.experimental_rerun()

st.sidebar.button("🚪 Sair", on_click=logout)
st.sidebar.write(f"Logado como: **{logged_user}**")

# ====================================================================================
#  A partir deste ponto, todo o código do app fica disponível somente após o login
# ====================================================================================

# CONSTANTES (os arquivos .xlsx devem estar na mesma pasta que este script)
EXCEL_PAGAR = "Contas a pagar 2025 Sistema.xlsx"
EXCEL_RECEBER = "Contas a Receber 2025 Sistema.xlsx"
ANEXOS_DIR = "anexos"

# ===============================
# FUNÇÕES AUXILIARES
# ===============================

def get_sheet_list(excel_path: str):
    """Retorna lista de abas, ignorando aba 'Tutorial' se existir."""
    try:
        wb = pd.ExcelFile(excel_path)
        return [s for s in wb.sheet_names if s.lower() != "tutorial"]
    except Exception:
        return []

def find_header_row(excel_path: str, sheet_name: str) -> int:
    """
    Retorna o índice da linha onde aparece 'Vencimento' no cabeçalho.
    """
    df_raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    for i, row in df_raw.iterrows():
        if any(str(cell).strip().lower() == "vencimento" for cell in row):
            return i
    return 0

def load_data(excel_path: str, sheet_name: str) -> pd.DataFrame:
    """
    Carrega a aba, detecta header, renomeia colunas e calcula status_pagamento.
    """
    header_row = find_header_row(excel_path, sheet_name)
    df = pd.read_excel(excel_path, sheet_name=sheet_name, skiprows=header_row, header=0)

    # Mapear nomes originais para nomes internos
    rename_map = {}
    for col in df.columns:
        nome = str(col).strip().lower()
        if nome == "data documento":
            rename_map[col] = "data_nf"
        elif nome == "descrição":
            rename_map[col] = "forma_pagamento"
        elif nome == "fornecedor":
            rename_map[col] = "fornecedor"
        elif nome == "documento":
            rename_map[col] = "os"
        elif nome == "vencimento":
            rename_map[col] = "vencimento"
        elif nome == "valor":
            rename_map[col] = "valor"
        elif nome == "estado":
            rename_map[col] = "estado"
        elif nome == "situação":
            rename_map[col] = "situacao"
        elif nome == "comprovante":
            rename_map[col] = "comprovante"
        elif nome == "boleto":
            rename_map[col] = "boleto"

    df = df.rename(columns=rename_map)

    expected_cols = {
        "data_nf", "forma_pagamento", "fornecedor", "os",
        "vencimento", "valor", "estado", "situacao", "boleto", "comprovante"
    }
    extra_cols = [c for c in df.columns if c not in expected_cols]
    if extra_cols:
        df = df.drop(extra_cols, axis=1)

    df = df.dropna(subset=["fornecedor", "valor"]).reset_index(drop=True)
    df["vencimento"] = pd.to_datetime(df["vencimento"], errors="coerce")
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce")

    # Calcula status_pagamento
    status_list = []
    hoje = datetime.now().date()
    for _, row in df.iterrows():
        pago = False
        if sheet_name.lower().startswith("contas a pagar"):
            if str(row.get("estado", "")).strip().lower() == "pago":
                pago = True
        else:
            if str(row.get("estado", "")).strip().lower() == "recebido":
                pago = True

        data_venc = row["vencimento"].date() if not pd.isna(row["vencimento"]) else None
        if pago:
            status_list.append("Em Dia")
        else:
            if data_venc:
                if data_venc < hoje:
                    status_list.append("Em Atraso")
                else:
                    status_list.append("A Vencer")
            else:
                status_list.append("Sem Data")

    df["status_pagamento"] = status_list
    return df

def rename_col_index(ws, target_name: str) -> int:
    """
    Retorna índice (1-based) da coluna cujo cabeçalho exato corresponde a target_name.
    """
    for row in ws.iter_rows(min_row=1, max_row=100, min_col=1, max_col=ws.max_column):
        for cell in row:
            if cell.value and str(cell.value).strip().lower() == target_name.lower():
                return cell.column
    defaults = {"vencimento": 5, "valor": 6, "estado": 7, "situação": 8}
    return defaults.get(target_name.lower(), 1)

def save_data(excel_path: str, sheet_name: str, df: pd.DataFrame):
    """
    Salva colunas 'valor', 'estado', 'situacao' e 'vencimento' de volta na planilha.
    """
    header_row = find_header_row(excel_path, sheet_name)
    wb = load_workbook(excel_path)
    ws = wb[sheet_name]

    for i, row in df.iterrows():
        excel_row = header_row + 1 + i
        ws.cell(row=excel_row + 1, column=rename_col_index(ws, "Valor"), value=row["valor"])
        ws.cell(row=excel_row + 1, column=rename_col_index(ws, "Estado"), value=row["estado"])
        ws.cell(row=excel_row + 1, column=rename_col_index(ws, "Situação"), value=row["situacao"])
        if pd.isna(row["vencimento"]):
            ws.cell(row=excel_row + 1, column=rename_col_index(ws, "Vencimento"), value=None)
        else:
            ws.cell(row=excel_row + 1, column=rename_col_index(ws, "Vencimento"), value=row["vencimento"])
    wb.save(excel_path)

def add_record(excel_path: str, sheet_name: str, record: dict):
    """
    Adiciona novo registro na próxima linha disponível da aba.
    """
    wb = load_workbook(excel_path)
    ws = wb[sheet_name]
    next_row = ws.max_row + 1

    valores = [
        record.get("data_nf", ""),
        record.get("forma_pagamento", ""),
        record.get("fornecedor", ""),
        record.get("os", ""),
        record.get("vencimento", ""),
        record.get("valor", ""),
        record.get("estado", ""),
        record.get("situacao", ""),
        record.get("boleto", ""),
        record.get("comprovante", "")
    ]
    for col_idx, val in enumerate(valores, start=1):
        ws.cell(row=next_row, column=col_idx, value=val)

    wb.save(excel_path)

# Garante pasta de anexos
for pasta in ["Contas a Pagar", "Contas a Receber"]:
    os.makedirs(os.path.join(ANEXOS_DIR, pasta), exist_ok=True)

# ===============================
# LÓGICA DO STREAMLIT
# ===============================
st.sidebar.markdown(
    """
    ## 📂 Navegação  
    Selecione a seção desejada para visualizar e gerenciar  
    suas contas a pagar e receber.  
    """
)
page = st.sidebar.radio("", ["Dashboard", "Contas a Pagar", "Contas a Receber"], index=0)

# Cabeçalho principal
st.markdown("""
<div style="text-align: center; color: #4B8BBE; margin-bottom: 10px;">
    <h1>💼 Sistema Financeiro 2025</h1>
    <p style="color: #555; font-size: 16px;">Dashboard avançado com estatísticas e gráficos interativos.</p>
</div>
""", unsafe_allow_html=True)
st.markdown("---")

# ------------------------
#  SEÇÃO: DASHBOARD
# ------------------------
if page == "Dashboard":
    st.subheader("📊 Painel de Controle Financeiro Avançado")
    sheets_p = get_sheet_list(EXCEL_PAGAR)
    sheets_r = get_sheet_list(EXCEL_RECEBER)

    tabs = st.tabs(["📥 Contas a Pagar", "📤 Contas a Receber"])

    # --------------------------------
    # CONTAS A PAGAR (Aba 1)
    # --------------------------------
    with tabs[0]:
        if not sheets_p:
            st.warning("Nenhuma aba encontrada em 'Contas a Pagar'.")
        else:
            df_all_p = pd.concat([load_data(EXCEL_PAGAR, s) for s in sheets_p], ignore_index=True)

            total_p = df_all_p["valor"].sum()
            num_lanc_p = len(df_all_p)
            media_p = df_all_p["valor"].mean() if num_lanc_p else 0
            atrasados_p = df_all_p[df_all_p["status_pagamento"] == "Em Atraso"]
            num_atras_p = len(atrasados_p)
            perc_atras_p = (num_atras_p / num_lanc_p * 100) if num_lanc_p else 0

            status_counts_p = (
                df_all_p["status_pagamento"]
                .value_counts()
                .rename_axis("status")
                .reset_index(name="contagem")
            )

            st.markdown(
                "<div style='padding:10px; background-color:#E8F8F5; border-radius:8px;'>"
                "<strong>Contas a Pagar - Estatísticas Gerais</strong></div>",
                unsafe_allow_html=True
            )
            c1, c2, c3, c4, c5 = st.columns([1.5, 1.5, 1.5, 1.5, 2])
            c1.metric("Total a Pagar", f"R$ {total_p:,.2f}")
            c2.metric("Nº Lançamentos", f"{num_lanc_p}")
            c3.metric("Média Valores", f"R$ {media_p:,.2f}")
            c4.metric("Em Atraso (%)", f"{perc_atras_p:.1f}% ({num_atras_p})")
            with c5:
                st.markdown("##### Distribuição por Status")
                st.bar_chart(status_counts_p.set_index("status")["contagem"])

            st.markdown("---")
            st.markdown("#### 📈 Evolução Mensal de Gastos")
            df_all_p["mes_ano"] = df_all_p["vencimento"].dt.to_period("M")
            monthly_group_p = (
                df_all_p
                .groupby("mes_ano")
                .agg(
                    total_mes=("valor", "sum"),
                    pagos_mes=("valor", lambda x: x[df_all_p.loc[x.index, "status_pagamento"] == "Em Dia"].sum()),
                    pendentes_mes=("valor", lambda x: x[df_all_p.loc[x.index, "status_pagamento"] != "Em Dia"].sum())
                )
                .reset_index()
            )
            monthly_group_p["mes_ano_str"] = monthly_group_p["mes_ano"].dt.strftime("%b/%Y")
            monthly_group_p = monthly_group_p.set_index("mes_ano_str")

            st.line_chart(monthly_group_p[["total_mes", "pagos_mes", "pendentes_mes"]])

            st.markdown("---")
            st.subheader("💾 Exportar Planilhas Originais (Contas a Pagar)")
            ep1, ep2 = st.columns(2)
            with ep1:
                try:
                    with open(EXCEL_PAGAR, "rb") as f:
                        dados_p = f.read()
                    st.download_button(
                        label="Download Excel (Pagar)",
                        data=dados_p,
                        file_name=EXCEL_PAGAR,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except FileNotFoundError:
                    st.error(f"'{EXCEL_PAGAR}' não encontrado.")
            with ep2:
                st.info("Para detalhes, acesse 'Contas a Pagar' no menu lateral.")

    # --------------------------------
    # CONTAS A RECEBER (Aba 2)
    # --------------------------------
    with tabs[1]:
        if not sheets_r:
            st.warning("Nenhuma aba encontrada em 'Contas a Receber'.")
        else:
            df_all_r = pd.concat([load_data(EXCEL_RECEBER, s) for s in sheets_r], ignore_index=True)

            total_r = df_all_r["valor"].sum()
            num_lanc_r = len(df_all_r)
            media_r = df_all_r["valor"].mean() if num_lanc_r else 0
            atrasados_r = df_all_r[df_all_r["status_pagamento"] == "Em Atraso"]
            num_atras_r = len(atrasados_r)
            perc_atras_r = (num_atras_r / num_lanc_r * 100) if num_lanc_r else 0

            status_counts_r = (
                df_all_r["status_pagamento"]
                .value_counts()
                .rename_axis("status")
                .reset_index(name="contagem")
            )

            st.markdown(
                "<div style='padding:10px; background-color:#FEF9E7; border-radius:8px;'>"
                "<strong>Contas a Receber - Estatísticas Gerais</strong></div>",
                unsafe_allow_html=True
            )
            d1, d2, d3, d4, d5 = st.columns([1.5, 1.5, 1.5, 1.5, 2])
            d1.metric("Total a Receber", f"R$ {total_r:,.2f}")
            d2.metric("Nº Lançamentos", f"{num_lanc_r}")
            d3.metric("Média Valores", f"R$ {media_r:,.2f}")
            d4.metric("Em Atraso (%)", f"{perc_atras_r:.1f}% ({num_atras_r})")
            with d5:
                st.markdown("##### Distribuição por Status")
                st.bar_chart(status_counts_r.set_index("status")["contagem"])

            st.markdown("---")
            st.markdown("#### 📈 Evolução Mensal de Recebimentos")
            df_all_r["mes_ano"] = df_all_r["vencimento"].dt.to_period("M")
            monthly_group_r = (
                df_all_r
                .groupby("mes_ano")
                .agg(
                    total_mes=("valor", "sum"),
                    recebidos_mes=("valor", lambda x: x[df_all_r.loc[x.index, "status_pagamento"] == "Em Dia"].sum()),
                    pendentes_mes=("valor", lambda x: x[df_all_r.loc[x.index, "status_pagamento"] != "Em Dia"].sum())
                )
                .reset_index()
            )
            monthly_group_r["mes_ano_str"] = monthly_group_r["mes_ano"].dt.strftime("%b/%Y")
            monthly_group_r = monthly_group_r.set_index("mes_ano_str")

            st.line_chart(monthly_group_r[["total_mes", "recebidos_mes", "pendentes_mes"]])

            st.markdown("---")
            st.subheader("💾 Exportar Planilhas Originais (Contas a Receber)")
            er1, er2 = st.columns(2)
            with er1:
                try:
                    with open(EXCEL_RECEBER, "rb") as f:
                        dados_r = f.read()
                    st.download_button(
                        label="Download Excel (Receber)",
                        data=dados_r,
                        file_name=EXCEL_RECEBER,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except FileNotFoundError:
                    st.error(f"'{EXCEL_RECEBER}' não encontrado.")
            with er2:
                st.info("Para detalhes, acesse 'Contas a Receber' no menu lateral.")

# ------------------------
#  SEÇÃO: CONTAS A PAGAR
# ------------------------
elif page == "Contas a Pagar":
    st.subheader("🗂️ Contas a Pagar")
    sheets = get_sheet_list(EXCEL_PAGAR)
    if not sheets:
        st.error(f"'{EXCEL_PAGAR}' não encontrado ou sem abas válidas.")
        st.stop()

    aba = st.selectbox("Selecione o mês:", sheets, index=0)
    df = load_data(EXCEL_PAGAR, aba)

    if df.empty:
        st.info("Nenhum registro encontrado para este mês.")
    else:
        with st.expander("🔍 Filtros"):
            colf1, colf2 = st.columns(2)
            with colf1:
                fornec_list = df["fornecedor"].dropna().astype(str).unique().tolist()
                forn = st.selectbox("Fornecedor", ["Todos"] + sorted(fornec_list))
            with colf2:
                est_list = df["estado"].dropna().astype(str).unique().tolist()
                status_sel = st.selectbox("Estado/Status", ["Todos"] + sorted(est_list))

        if forn != "Todos":
            df = df[df["fornecedor"] == forn]
        if status_sel != "Todos":
            df = df[df["estado"] == status_sel]

        st.markdown("<hr style='border:1px solid #ddd;'>", unsafe_allow_html=True)

        if df.empty:
            st.warning("Nenhum registro corresponde aos filtros selecionados.")
        else:
            cols_esperadas = ["data_nf", "fornecedor", "valor", "vencimento", "status_pagamento"]
            cols_para_exibir = [c for c in cols_esperadas if c in df.columns]
            st.markdown("#### 📋 Lista de Lançamentos")
            st.dataframe(df[cols_para_exibir], height=250)
            st.markdown("---")

            with st.expander("✏️ Editar Registro"):
                idx = st.number_input("Índice da linha:", min_value=0, max_value=len(df) - 1, step=1)
                rec = df.iloc[idx]

                colv1, colv2 = st.columns(2)
                with colv1:
                    new_val = st.number_input("Valor:", value=float(rec["valor"]), key="valores")
                    default_dt = rec["vencimento"].date() if pd.notna(rec["vencimento"]) else date.today()
                    new_venc = st.date_input("Vencimento:", value=default_dt, key="vencimento")
                with colv2:
                    estado_uni = df["estado"].dropna().astype(str).unique().tolist()
                    try:
                        est_idx = estado_uni.index(str(rec["estado"]))
                    except ValueError:
                        est_idx = 0
                    new_estado = st.selectbox("Estado:", options=estado_uni, index=est_idx, key="estado")

                    situ_uni = df["situacao"].dropna().astype(str).unique().tolist()
                    try:
                        sit_idx = situ_uni.index(str(rec["situacao"]))
                    except ValueError:
                        sit_idx = 0
                    new_sit = st.selectbox("Situação:", options=situ_uni, index=sit_idx, key="situacao")

                if st.button("💾 Salvar Alterações"):
                    df.loc[df.index[idx], ["valor", "vencimento", "estado", "situacao"]] = [
                        new_val, pd.to_datetime(new_venc), new_estado, new_sit
                    ]
                    save_data(EXCEL_PAGAR, aba, df)
                    st.success("Registro atualizado com sucesso!")

            st.markdown("---")

            with st.expander("📎 Anexar Documentos"):
                idx2 = st.number_input(
                    "Índice para anexar:", min_value=0, max_value=len(df) - 1, step=1, key="idx_anex"
                )
                uploaded = st.file_uploader(
                    "Selecione (pdf/jpg/png):", type=["pdf", "jpg", "png"], key=f"up_pagar_{aba}_{idx2}"
                )
                if uploaded:
                    destino = os.path.join(
                        ANEXOS_DIR, "Contas a Pagar", f"Pagar_{aba}_{idx2}_{uploaded.name}"
                    )
                    with open(destino, "wb") as f:
                        f.write(uploaded.getbuffer())
                    st.success(f"Documento salvo em: {destino}")

            st.markdown("---")

            with st.expander("➕ Adicionar Nova Conta"):
                coln1, coln2 = st.columns(2)
                with coln1:
                    data_nf = st.date_input("Data N/F:", value=date.today())
                    forma_pag = st.text_input("Descrição:")
                    forn_new = st.text_input("Fornecedor:")
                with coln2:
                    os_new = st.text_input("Documento/OS:")
                    venc_new = st.date_input("Data de Vencimento:", value=date.today())
                    valor_new = st.number_input("Valor (R$):", min_value=0.0, format="%.2f")

                estado_opt = ["Em Aberto", "Pago"]
                situ_opt = ["Em Atraso", "Pago", "Em Aberto"]
                estado_new = st.selectbox("Estado:", options=estado_opt)
                situ_new = st.selectbox("Situação:", options=situ_opt)
                boleto_file = st.file_uploader(
                    "Boleto (opcional):", type=["pdf", "jpg", "png"], key="boleto_pagar"
                )
                comprov_file = st.file_uploader(
                    "Comprovante (opcional):", type=["pdf", "jpg", "png"], key="comprov_pagar"
                )
                if st.button("➕ Adicionar Conta"):
                    boleto_path = ""
                    comprov_path = ""
                    if boleto_file:
                        boleto_path = os.path.join(
                            ANEXOS_DIR, "Contas a Pagar", f"Pagar_{aba}_boleto_{boleto_file.name}"
                        )
                        with open(boleto_path, "wb") as fb:
                            fb.write(boleto_file.getbuffer())
                    if comprov_file:
                        comprov_path = os.path.join(
                            ANEXOS_DIR, "Contas a Pagar", f"Pagar_{aba}_comprov_{comprov_file.name}"
                        )
                        with open(comprov_path, "wb") as fc:
                            fc.write(comprov_file.getbuffer())

                    record = {
                        "data_nf": data_nf,
                        "forma_pagamento": forma_pag,
                        "fornecedor": forn_new,
                        "os": os_new,
                        "vencimento": venc_new,
                        "valor": valor_new,
                        "estado": estado_new,
                        "situacao": situ_new,
                        "boleto": boleto_path,
                        "comprovante": comprov_path,
                    }
                    add_record(EXCEL_PAGAR, aba, record)
                    st.success("Nova conta adicionada com sucesso!")

            st.markdown("---")

            st.subheader("💾 Exportar Aba Atual")
            try:
                save_data(EXCEL_PAGAR, aba, df)
                with open(EXCEL_PAGAR, "rb") as fx:
                    bytes_data = fx.read()
                st.download_button(
                    label=f"Exportar '{aba}'",
                    data=bytes_data,
                    file_name=f"Contas a Pagar - {aba}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"Erro ao preparar download: {e}")

# ------------------------------------
#  SEÇÃO: CONTAS A RECEBER
# ------------------------------------
elif page == "Contas a Receber":
    st.subheader("🗂️ Contas a Receber")
    sheets = get_sheet_list(EXCEL_RECEBER)
    if not sheets:
        st.error(f"'{EXCEL_RECEBER}' não encontrado ou sem abas válidas.")
        st.stop()

    aba = st.selectbox("Selecione o mês:", sheets, index=0)
    df = load_data(EXCEL_RECEBER, aba)

    if df.empty:
        st.info("Nenhum registro encontrado para este mês.")
    else:
        with st.expander("🔍 Filtros"):
            colf1, colf2 = st.columns(2)
            with colf1:
                fornec_list = df["fornecedor"].dropna().astype(str).unique().tolist()
                forn = st.selectbox("Fornecedor", ["Todos"] + sorted(fornec_list))
            with colf2:
                est_list = df["estado"].dropna().astype(str).unique().tolist()
                status_sel = st.selectbox("Estado/Status", ["Todos"] + sorted(est_list))

        if forn != "Todos":
            df = df[df["fornecedor"] == forn]
        if status_sel != "Todos":
            df = df[df["estado"] == status_sel]

        st.markdown("<hr style='border:1px solid #ddd;'>", unsafe_allow_html=True)

        if df.empty:
            st.warning("Nenhum registro corresponde aos filtros selecionados.")
        else:
            cols_esperadas = ["data_nf", "fornecedor", "valor", "vencimento", "status_pagamento"]
            cols_para_exibir = [c for c in cols_esperadas if c in df.columns]
            st.markdown("#### 📋 Lista de Lançamentos")
            st.dataframe(df[cols_para_exibir], height=250)
            st.markdown("---")

            with st.expander("✏️ Editar Registro"):
                idx = st.number_input(
                    "Índice da linha:", min_value=0, max_value=len(df) - 1, step=1, key="idx_rec_r"
                )
                rec = df.iloc[idx]

                colv1, colv2 = st.columns(2)
                with colv1:
                    new_val = st.number_input("Valor:", value=float(rec["valor"]), key="valores_r")
                    default_dt = rec["vencimento"].date() if pd.notna(rec["vencimento"]) else date.today()
                    new_venc = st.date_input("Vencimento:", value=default_dt, key="vencimento_r")
                with colv2:
                    estado_uni = df["estado"].dropna().astype(str).unique().tolist()
                    try:
                        est_idx = estado_uni.index(str(rec["estado"]))
                    except ValueError:
                        est_idx = 0
                    new_estado = st.selectbox("Estado:", options=estado_uni, index=est_idx, key="estado_r")

                    situ_uni = df["situacao"].dropna().astype(str).unique().tolist()
                    try:
                        sit_idx = situ_uni.index(str(rec["situacao"]))
                    except ValueError:
                        sit_idx = 0
                    new_sit = st.selectbox("Situação:", options=situ_uni, index=sit_idx, key="situacao_r")

                if st.button("💾 Salvar Alterações", key="salvar_r"):
                    df.loc[df.index[idx], ["valor", "vencimento", "estado", "situacao"]] = [
                        new_val,
                        pd.to_datetime(new_venc),
                        new_estado,
                        new_sit,
                    ]
                    save_data(EXCEL_RECEBER, aba, df)
                    st.success("Registro atualizado com sucesso!")

            st.markdown("---")

            with st.expander("📎 Anexar Documentos"):
                idx2 = st.number_input(
                    "Índice para anexar:", min_value=0, max_value=len(df) - 1, step=1, key="idx_anex_r"
                )
                uploaded = st.file_uploader(
                    "Selecione (pdf/jpg/png):", type=["pdf", "jpg", "png"], key=f"up_receber_{aba}_{idx2}"
                )
                if uploaded:
                    destino = os.path.join(
                        ANEXOS_DIR, "Contas a Receber", f"Receber_{aba}_{idx2}_{uploaded.name}"
                    )
                    with open(destino, "wb") as f:
                        f.write(uploaded.getbuffer())
                    st.success(f"Documento salvo em: {destino}")

            st.markdown("---")

            with st.expander("➕ Adicionar Nova Conta"):
                coln1, coln2 = st.columns(2)
                with coln1:
                    data_nf = st.date_input("Data N/F:", value=date.today(), key="data_nf_r")
                    forma_pag = st.text_input("Descrição:", key="forma_pag_r")
                    forn_new = st.text_input("Fornecedor:", key="forn_r")
                with coln2:
                    os_new = st.text_input("Documento/OS:", key="os_r")
                    venc_new = st.date_input("Data de Vencimento:", value=date.today(), key="venc_new_r")
                    valor_new = st.number_input("Valor (R$):", min_value=0.0, format="%.2f", key="valor_new_r")

                estado_opt = ["A Receber", "Recebido"]
                situ_opt = ["Em Atraso", "Recebido", "A Receber"]
                estado_new = st.selectbox("Estado:", options=estado_opt, key="estado_new_r")
                situ_new = st.selectbox("Situação:", options=situ_opt, key="situ_new_r")
                boleto_file = st.file_uploader("Boleto (opcional):", type=["pdf", "jpg", "png"], key="boleto_r")
                comprov_file = st.file_uploader("Comprovante (opcional):", type=["pdf", "jpg", "png"], key="comprov_r")
                if st.button("➕ Adicionar Conta", key="add_r"):
                    boleto_path = ""
                    comprov_path = ""
                    if boleto_file:
                        boleto_path = os.path.join(
                            ANEXOS_DIR, "Contas a Receber", f"Receber_{aba}_boleto_{boleto_file.name}"
                        )
                        with open(boleto_path, "wb") as fb:
                            fb.write(boleto_file.getbuffer())
                    if comprov_file:
                        comprov_path = os.path.join(
                            ANEXOS_DIR, "Contas a Receber", f"Receber_{aba}_comprov_{comprov_file.name}"
                        )
                        with open(comprov_path, "wb") as fc:
                            fc.write(comprov_file.getbuffer())

                    record = {
                        "data_nf": data_nf,
                        "forma_pagamento": forma_pag,
                        "fornecedor": forn_new,
                        "os": os_new,
                        "vencimento": venc_new,
                        "valor": valor_new,
                        "estado": estado_new,
                        "situacao": situ_new,
                        "boleto": boleto_path,
                        "comprovante": comprov_path,
                    }
                    add_record(EXCEL_RECEBER, aba, record)
                    st.success("Nova conta adicionada com sucesso!")

            st.markdown("---")

            st.subheader("💾 Exportar Aba Atual")
            try:
                save_data(EXCEL_RECEBER, aba, df)
                with open(EXCEL_RECEBER, "rb") as fx:
                    bytes_data = fx.read()
                st.download_button(
                    label=f"Exportar '{aba}'",
                    data=bytes_data,
                    file_name=f"Contas a Receber - {aba}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"Erro ao preparar download: {e}")

# ===============================
#  RODAPÉ
# ===============================
st.markdown("""
<div style="text-align: center; font-size:12px; color:gray; margin-top: 20px;">
    <p>© 2025 Desenvolvido por Vinicius Magalhães</p>
</div>
""", unsafe_allow_html=True)
