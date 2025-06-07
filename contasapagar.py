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
    layout="wide"
)

# ====================================================================
#  Autenticação simples (sem bibliotecas externas), formulário centralizado
# ====================================================================
VALID_USERS = {
    "Vinicius": "vinicius4223",
    "Flavio": "1234",
}

def check_login(username: str, password: str) -> bool:
    return VALID_USERS.get(username) == password

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = ""

# Se não estiver logado, exibe apenas o formulário de login
if not st.session_state.logged_in:
    st.write("\n" * 5)  # puxa um pouco para baixo, para centralizar vertical

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.title("🔒 Login")
        username_input = st.text_input("Usuário:")
        password_input = st.text_input("Senha:", type="password")
        if st.button("Entrar"):
            if check_login(username_input, password_input):
                st.session_state.logged_in = True
                st.session_state.username = username_input
            else:
                st.error("Usuário ou senha inválidos.")
    st.stop()

# Usuário já autenticado
logged_user = st.session_state.username
st.sidebar.write(f"Logado como: **{logged_user}**")

# ====================================================================================
#  A partir deste ponto, todo o código do app fica disponível somente após o login
# ====================================================================================

EXCEL_PAGAR   = "Contas a pagar 2025 Sistema.xlsx"
EXCEL_RECEBER = "Contas a Receber 2025 Sistema.xlsx"
ANEXOS_DIR    = "anexos"
# Lista fixa de meses "01".."12", para permitir seleção de meses futuros mesmo que a aba ainda não exista
FULL_MONTHS   = [f"{i:02d}" for i in range(1, 13)]


# ===============================
# FUNÇÕES AUXILIARES
# ===============================

def get_existing_sheets(excel_path: str) -> list[str]:
    """
    Retorna as abas numéricas existentes no arquivo (ex: '01','02', ... '12'),
    ignorando 'Tutorial'. Se der erro ao abrir, retorna lista vazia.
    """
    try:
        wb = pd.ExcelFile(excel_path)
        numeric_sheets = []
        for s in wb.sheet_names:
            nome = s.strip()
            if nome.lower() == "tutorial":
                continue
            if nome.isdigit():
                numeric_sheets.append(nome)
        return sorted(numeric_sheets)
    except Exception:
        return []


def load_data(excel_path: str, sheet_name: str) -> pd.DataFrame:
    """
    Carrega dados da aba sheet_name (por exemplo, "06"). Se a aba não existir,
    retorna um DataFrame vazio com todas as colunas esperadas.

    Usa skiprows=7 porque o cabeçalho real (com "Vencimento", "Valor" etc.) está na linha 8 do Excel.
    """
    cols = [
        "data_nf", "forma_pagamento", "fornecedor", "os",
        "vencimento", "valor", "estado", "situacao", "boleto", "comprovante"
    ]

    # Se o arquivo não existe, devolve DF vazio com colunas
    if not os.path.isfile(excel_path):
        df_empty = pd.DataFrame(columns=cols + ["status_pagamento"])
        return df_empty

    existing = get_existing_sheets(excel_path)
    if sheet_name not in existing:
        # Aba ainda não existe → DF vazio
        df_empty = pd.DataFrame(columns=cols + ["status_pagamento"])
        return df_empty

    # Lê sempre pulando as primeiras 7 linhas, pois o cabeçalho real inicia na 8ª linha
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, skiprows=7, header=0)
    except Exception:
        # Se der erro lendo (formato inesperado), retorna DF vazio
        df_empty = pd.DataFrame(columns=cols + ["status_pagamento"])
        return df_empty

    # Mapeia colunas do Excel para nossos nomes internos
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

    # Mantém apenas as colunas esperadas (descarta extras)
    expected_cols = set(cols)
    extras = [c for c in df.columns if c not in expected_cols]
    if extras:
        df = df.drop(extras, axis=1)

    # Remove linhas sem fornecedor ou valor
    df = df.dropna(subset=["fornecedor", "valor"]).reset_index(drop=True)

    # Conversões de tipo
    df["vencimento"] = pd.to_datetime(df["vencimento"], errors="coerce")
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce")

    # Calcula coluna de status_pagamento: se estado == "Pago", status="Pago";
    # senão, se vencimento < hoje → "Em Atraso"; senão → "A Vencer"; ou "Sem Data".
    status_list = []
    hoje = datetime.now().date()
    for _, row in df.iterrows():
        estado_atual = str(row.get("estado", "")).strip().lower()
        if estado_atual == "pago":
            status_list.append("Pago")
            continue
        data_venc = row["vencimento"].date() if pd.notna(row["vencimento"]) else None
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
    Dado um worksheet (ws), retorna o índice (1-based) da coluna cujo cabeçalho
    bate exatamente (case-insensitive) com target_name. Se não achar,
    retorna valor padrão: Vencimento=5, Valor=6, Estado=7, Situação=8.
    """
    for row in ws.iter_rows(min_row=1, max_row=100, min_col=1, max_col=ws.max_column):
        for cell in row:
            if cell.value and str(cell.value).strip().lower() == target_name.lower():
                return cell.column
    defaults = {"vencimento": 5, "valor": 6, "estado": 7, "situação": 8}
    return defaults.get(target_name.lower(), 1)


def save_data(excel_path: str, sheet_name: str, df: pd.DataFrame):
    """
    Salva de volta no Excel apenas as colunas 'valor', 'estado', 'situacao' e 'vencimento'
    na aba sheet_name, mantendo cabeçalhos e fórmulas originais.
    """
    wb = load_workbook(excel_path)
    ws = wb[sheet_name]
    # Cabeçalho está na linha 8, então a primeira linha de dados é 9 (índice 8 0-based).
    # Porém, como pulamos 7 linhas no load_data, basta usar i+8.
    for i, row in df.iterrows():
        excel_row = i + 8  # 0-based i → linha real = i+8+1 1-based
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
    Adiciona um novo registro na próxima linha disponível da aba sheet_name.
    Se a aba não existir, cria-a automaticamente duplicando a primeira aba numérica válida.
    Grava: data_nf, forma_pagamento, fornecedor, os, vencimento, valor, estado, situacao, boleto, comprovante.
    """
    wb = load_workbook(excel_path)
    existing = [s.strip() for s in wb.sheetnames]

    if sheet_name not in existing:
        # Cria nova aba a partir da primeira aba numérica existente
        numeric = [s for s in existing if s.isdigit()]
        if numeric:
            template_ws = wb[numeric[0]]
        else:
            template_ws = wb[wb.sheetnames[0]]
        new_ws = wb.copy_worksheet(template_ws)
        new_ws.title = sheet_name
        ws = new_ws
    else:
        ws = wb[sheet_name]

    next_row = ws.max_row + 1
    vals = [
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
    for col_idx, val in enumerate(vals, start=1):
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

    # Verifica existência dos arquivos
    if not os.path.isfile(EXCEL_PAGAR):
        st.error(f"Arquivo '{EXCEL_PAGAR}' não encontrado. Verifique o caminho.")
        st.stop()
    if not os.path.isfile(EXCEL_RECEBER):
        st.error(f"Arquivo '{EXCEL_RECEBER}' não encontrado. Verifique o caminho.")
        st.stop()

    sheets_p = get_existing_sheets(EXCEL_PAGAR)
    sheets_r = get_existing_sheets(EXCEL_RECEBER)

    tabs = st.tabs(["📥 Contas a Pagar", "📤 Contas a Receber"])

    # ------------------------
    # CONTAS A PAGAR (Aba 1)
    # ------------------------
    with tabs[0]:
        if not sheets_p:
            st.warning("'Contas a Pagar' encontrado, mas não há abas numéricas válidas (espera-se '01'..'12').")
        else:
            df_all_p = pd.concat([load_data(EXCEL_PAGAR, s) for s in sheets_p], ignore_index=True)
            total_p      = df_all_p["valor"].sum()
            num_lanc_p   = len(df_all_p)
            media_p      = df_all_p["valor"].mean() if num_lanc_p else 0
            atrasados_p  = df_all_p[df_all_p["status_pagamento"] == "Em Atraso"]
            num_atras_p  = len(atrasados_p)
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
            c1.metric("Total a Pagar",   f"R$ {total_p:,.2f}")
            c2.metric("Nº Lançamentos",   f"{num_lanc_p}")
            c3.metric("Média Valores",    f"R$ {media_p:,.2f}")
            c4.metric("Em Atraso (%)",    f"{perc_atras_p:.1f}% ({num_atras_p})")
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
                    pagos_mes=("valor", lambda x: x[df_all_p.loc[x.index, "status_pagamento"] == "Pago"].sum()),
                    pendentes_mes=("valor", lambda x: x[df_all_p.loc[x.index, "status_pagamento"] != "Pago"].sum())
                )
                .reset_index()
            )
            monthly_group_p["mes_ano_str"] = monthly_group_p["mes_ano"].dt.strftime("%b/%Y")
            monthly_group_p = monthly_group_p.set_index("mes_ano_str")
            st.line_chart(monthly_group_p[["total_mes", "pagos_mes", "pendentes_mes"]])

            st.markdown("---")

            st.markdown("#### 📊 Percentual por Status de Pagamento")
            status_counts_p["percentual"] = status_counts_p["contagem"] / num_lanc_p * 100
            df_status_pct = status_counts_p.set_index("status")[["percentual"]]
            df_status_pct.columns = ["% (%)"]
            st.bar_chart(df_status_pct)

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

    # ------------------------
    # CONTAS A RECEBER (Aba 2)
    # ------------------------
    with tabs[1]:
        if not sheets_r:
            st.warning("'Contas a Receber' encontrado, mas não há abas numéricas válidas (espera-se '01'..'12').")
        else:
            df_all_r = pd.concat([load_data(EXCEL_RECEBER, s) for s in sheets_r], ignore_index=True)
            total_r      = df_all_r["valor"].sum()
            num_lanc_r   = len(df_all_r)
            media_r      = df_all_r["valor"].mean() if num_lanc_r else 0
            atrasados_r  = df_all_r[df_all_r["status_pagamento"] == "Em Atraso"]
            num_atras_r  = len(atrasados_r)
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
            d1.metric("Total a Receber",   f"R$ {total_r:,.2f}")
            d2.metric("Nº Lançamentos",   f"{num_lanc_r}")
            d3.metric("Média Valores",    f"R$ {media_r:,.2f}")
            d4.metric("Em Atraso (%)",    f"{perc_atras_r:.1f}% ({num_atras_r})")
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
                    recebidos_mes=("valor", lambda x: x[df_all_r.loc[x.index, "status_pagamento"] == "Pago"].sum()),
                    pendentes_mes=("valor", lambda x: x[df_all_r.loc[x.index, "status_pagamento"] != "Pago"].sum())
                )
                .reset_index()
            )
            monthly_group_r["mes_ano_str"] = monthly_group_r["mes_ano"].dt.strftime("%b/%Y")
            monthly_group_r = monthly_group_r.set_index("mes_ano_str")
            st.line_chart(monthly_group_r[["total_mes", "recebidos_mes", "pendentes_mes"]])

            st.markdown("---")

            st.markdown("#### 📊 Percentual por Status de Recebimento")
            status_counts_r["percentual"] = status_counts_r["contagem"] / num_lanc_r * 100
            df_status_pct_r = status_counts_r.set_index("status")[["percentual"]]
            df_status_pct_r.columns = ["% (%)"]
            st.bar_chart(df_status_pct_r)

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

    if not os.path.isfile(EXCEL_PAGAR):
        st.error(f"Arquivo '{EXCEL_PAGAR}' não encontrado. Verifique o caminho.")
        st.stop()

    existing = get_existing_sheets(EXCEL_PAGAR)

    # Seleção de mês: sempre mostra "01".."12", mesmo que ainda não exista no Excel
    aba = st.selectbox("Selecione o mês:", FULL_MONTHS, index=0)
    df = load_data(EXCEL_PAGAR, aba)  # se a aba não existir, retorna DF vazio

    if df.empty:
        st.info("Nenhum registro encontrado para este mês (ou a aba não existia).")

    view_sel = st.radio("Visualizar:", ["Todos", "Pagas", "Pendentes"], horizontal=True)
    if view_sel == "Pagas":
        df_display = df[df["estado"].str.strip().str.lower() == "pago"].copy()
    elif view_sel == "Pendentes":
        df_display = df[df["estado"].str.strip().str.lower() != "pago"].copy()
    else:
        df_display = df.copy()

    with st.expander("🔍 Filtros"):
        colf1, colf2 = st.columns(2)
        with colf1:
            fornec_list = df["fornecedor"].dropna().astype(str).unique().tolist()
            forn = st.selectbox("Fornecedor", ["Todos"] + sorted(fornec_list))
        with colf2:
            est_list = df["estado"].dropna().astype(str).unique().tolist()
            status_sel = st.selectbox("Estado/Status", ["Todos"] + sorted(est_list))

    if "forn" in locals() and forn != "Todos":
        df_display = df_display[df_display["fornecedor"] == forn]
    if "status_sel" in locals() and status_sel != "Todos":
        df_display = df_display[df_display["estado"] == status_sel]

    st.markdown("<hr style='border:1px solid #ddd;'>", unsafe_allow_html=True)

    if df_display.empty:
        st.warning("Nenhum registro para os filtros/visualização selecionados.")
    else:
        cols_esperadas = ["data_nf", "fornecedor", "valor", "vencimento", "estado", "status_pagamento"]
        cols_para_exibir = [c for c in cols_esperadas if c in df_display.columns]

        st.markdown("#### 📋 Lista de Lançamentos")
        table_placeholder = st.empty()
        table_placeholder.dataframe(df_display[cols_para_exibir], height=250)

    st.markdown("---")

    # ======= EDIÇÃO DE REGISTRO =======
    with st.expander("✏️ Editar Registro"):
        idx = st.number_input(
            "Índice da linha (baseado na lista acima):",
            min_value=0, max_value=len(df_display) - 1 if not df_display.empty else 0, step=1, key="edit_pagar"
        )
        if not df_display.empty:
            rec = df_display.iloc[idx]
            # Localiza índice original no DF completo
            orig_idx_candidates = df[
                (df["fornecedor"] == rec["fornecedor"]) &
                (df["valor"] == rec["valor"]) &
                (df["vencimento"] == rec["vencimento"])
            ].index
            orig_idx = orig_idx_candidates[0] if len(orig_idx_candidates) > 0 else rec.name

            colv1, colv2 = st.columns(2)
            with colv1:
                new_val = st.number_input("Valor:", value=float(rec["valor"]), key="novo_valor_pagar")
                default_dt = rec["vencimento"].date() if pd.notna(rec["vencimento"]) else date.today()
                new_venc = st.date_input("Vencimento:", value=default_dt, key="novo_vencimento_pagar")
            with colv2:
                estado_uni = df["estado"].dropna().astype(str).unique().tolist()
                try:
                    est_idx = estado_uni.index(str(rec["estado"]))
                except ValueError:
                    est_idx = 0
                new_estado = st.selectbox("Estado:", options=estado_uni, index=est_idx, key="novo_estado_pagar")

                situ_uni = df["situacao"].dropna().astype(str).unique().tolist()
                try:
                    sit_idx = situ_uni.index(str(rec["situacao"]))
                except ValueError:
                    sit_idx = 0
                new_sit = st.selectbox("Situação:", options=situ_uni, index=sit_idx, key="nova_situacao_pagar")

            if st.button("💾 Salvar Alterações", key="salvar_pagar"):
                df.at[orig_idx, "valor"] = new_val
                df.at[orig_idx, "vencimento"] = pd.to_datetime(new_venc)
                df.at[orig_idx, "estado"] = new_estado
                df.at[orig_idx, "situacao"] = new_sit

                save_data(EXCEL_PAGAR, aba, df)
                df = load_data(EXCEL_PAGAR, aba)
                st.success("Registro atualizado com sucesso!")

                if view_sel == "Pagas":
                    df_display = df[df["estado"].str.strip().str.lower() == "pago"].copy()
                elif view_sel == "Pendentes":
                    df_display = df[df["estado"].str.strip().str.lower() != "pago"].copy()
                else:
                    df_display = df.copy()

                if "forn" in locals() and forn != "Todos":
                    df_display = df_display[df_display["fornecedor"] == forn]
                if "status_sel" in locals() and status_sel != "Todos":
                    df_display = df_display[df_display["estado"] == status_sel]

                table_placeholder.dataframe(df_display[cols_para_exibir], height=250)

    st.markdown("---")

    # ======= ANEXAR DOCUMENTOS =======
    with st.expander("📎 Anexar Documentos"):
        if not df_display.empty:
            idx2 = st.number_input(
                "Índice para anexar (baseado na lista acima):",
                min_value=0, max_value=len(df_display) - 1, step=1, key="idx_anex_pagar"
            )
            rec_anex = df_display.iloc[idx2]
            orig_idx_anex_candidates = df[
                (df["fornecedor"] == rec_anex["fornecedor"]) &
                (df["valor"] == rec_anex["valor"]) &
                (df["vencimento"] == rec_anex["vencimento"])
            ].index
            orig_idx_anex = orig_idx_anex_candidates[0] if len(orig_idx_anex_candidates) > 0 else rec_anex.name

            uploaded = st.file_uploader(
                "Selecione (pdf/jpg/png):", type=["pdf", "jpg", "png"], key=f"up_pagar_{aba}_{idx2}"
            )
            if uploaded:
                destino = os.path.join(
                    ANEXOS_DIR, "Contas a Pagar", f"Pagar_{aba}_{orig_idx_anex}_{uploaded.name}"
                )
                with open(destino, "wb") as f:
                    f.write(uploaded.getbuffer())
                st.success(f"Documento salvo em: {destino}")

    st.markdown("---")

    # ======= ADICIONAR NOVA CONTA =======
    with st.expander("➕ Adicionar Nova Conta"):
        coln1, coln2 = st.columns(2)
        with coln1:
            data_nf   = st.date_input("Data N/F:", value=date.today(), key="nova_data_nf_pagar")
            forma_pag = st.text_input("Descrição:", key="nova_descricao_pagar")
            forn_new  = st.text_input("Fornecedor:", key="novo_fornecedor_pagar")
        with coln2:
            os_new    = st.text_input("Documento/OS:", key="novo_os_pagar")
            venc_new  = st.date_input("Data de Vencimento:", value=date.today(), key="novo_venc_pagar")
            valor_new = st.number_input("Valor (R$):", min_value=0.0, format="%.2f", key="novo_valor_pagar2")

        estado_opt   = ["Em Aberto", "Pago"]
        situ_opt     = ["Em Atraso", "Pago", "Em Aberto"]
        estado_new   = st.selectbox("Estado:", options=estado_opt, key="estado_novo_pagar")
        situ_new     = st.selectbox("Situação:", options=situ_opt,   key="situacao_novo_pagar")
        boleto_file   = st.file_uploader("Boleto (opcional):",   type=["pdf", "jpg", "png"], key="boleto_novo_pagar")
        comprov_file = st.file_uploader("Comprovante (opcional):", type=["pdf", "jpg", "png"], key="comprov_novo_pagar")

        if st.button("➕ Adicionar Conta", key="adicionar_pagar"):
            record = {
                "data_nf": data_nf,
                "forma_pagamento": forma_pag,
                "fornecedor": forn_new,
                "os": os_new,
                "vencimento": venc_new,
                "valor": valor_new,
                "estado": estado_new,
                "situacao": situ_new,
                "boleto": "",
                "comprovante": ""
            }
            if boleto_file:
                boleto_path = os.path.join(
                    ANEXOS_DIR, "Contas a Pagar", f"Pagar_{aba}_boleto_{boleto_file.name}"
                )
                with open(boleto_path, "wb") as fb:
                    fb.write(boleto_file.getbuffer())
                record["boleto"] = boleto_path
            if comprov_file:
                comprov_path = os.path.join(
                    ANEXOS_DIR, "Contas a Pagar", f"Pagar_{aba}_comprov_{comprov_file.name}"
                )
                with open(comprov_path, "wb") as fc:
                    fc.write(comprov_file.getbuffer())
                record["comprovante"] = comprov_path

            # Grava no Excel (cria aba "06" automaticamente, se necessário)
            add_record(EXCEL_PAGAR, aba, record)
            st.success("Nova conta adicionada com sucesso!")

            df = load_data(EXCEL_PAGAR, aba)
            if view_sel == "Pagas":
                df_display = df[df["estado"].str.strip().str.lower() == "pago"].copy()
            elif view_sel == "Pendentes":
                df_display = df[df["estado"].str.strip().str.lower() != "pago"].copy()
            else:
                df_display = df.copy()

            if "forn" in locals() and forn != "Todos":
                df_display = df_display[df_display["fornecedor"] == forn]
            if "status_sel" in locals() and status_sel != "Todos":
                df_display = df_display[df_display["estado"] == status_sel]

            table_placeholder.dataframe(df_display[cols_para_exibir], height=250)

    st.markdown("---")

    # ======= EXPORTAR ABA ATUAL =======
    st.subheader("💾 Exportar Aba Atual")
    try:
        df_to_save = load_data(EXCEL_PAGAR, aba)
        if not df_to_save.empty:
            save_data(EXCEL_PAGAR, aba, df_to_save)
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


# ------------------------
#  SEÇÃO: CONTAS A RECEBER
# ------------------------
elif page == "Contas a Receber":
    st.subheader("🗂️ Contas a Receber")

    if not os.path.isfile(EXCEL_RECEBER):
        st.error(f"Arquivo '{EXCEL_RECEBER}' não encontrado. Verifique o caminho.")
        st.stop()

    aba = st.selectbox("Selecione o mês:", FULL_MONTHS, index=0)
    df = load_data(EXCEL_RECEBER, aba)  # se a aba não existir, retorna DF vazio

    if df.empty:
        st.info("Nenhum registro encontrado para este mês (ou a aba não existia).")

    view_sel = st.radio("Visualizar:", ["Todos", "Recebidas", "Pendentes"], horizontal=True)
    if view_sel == "Recebidas":
        df_display = df[df["estado"].str.strip().str.lower() == "recebido"].copy()
    elif view_sel == "Pendentes":
        df_display = df[df["estado"].str.strip().str.lower() != "recebido"].copy()
    else:
        df_display = df.copy()

    with st.expander("🔍 Filtros"):
        colf1, colf2 = st.columns(2)
        with colf1:
            fornec_list = df["fornecedor"].dropna().astype(str).unique().tolist()
            forn = st.selectbox("Fornecedor", ["Todos"] + sorted(fornec_list))
        with colf2:
            est_list = df["estado"].dropna().astype(str).unique().tolist()
            status_sel = st.selectbox("Estado/Status", ["Todos"] + sorted(est_list))

    if "forn" in locals() and forn != "Todos":
        df_display = df_display[df_display["fornecedor"] == forn]
    if "status_sel" in locals() and status_sel != "Todos":
        df_display = df_display[df_display["estado"] == status_sel]

    st.markdown("<hr style='border:1px solid #ddd;'>", unsafe_allow_html=True)

    if df_display.empty:
        st.warning("Nenhum registro para os filtros/visualização selecionados.")
    else:
        cols_esperadas = ["data_nf", "fornecedor", "valor", "vencimento", "estado", "status_pagamento"]
        cols_para_exibir = [c for c in cols_esperadas if c in df_display.columns]

        st.markdown("#### 📋 Lista de Lançamentos")
        table_placeholder_r = st.empty()
        table_placeholder_r.dataframe(df_display[cols_para_exibir], height=250)

    st.markdown("---")

    # ======= EDIÇÃO DE REGISTRO =======
    with st.expander("✏️ Editar Registro"):
        idx = st.number_input(
            "Índice da linha (baseado na lista acima):",
            min_value=0, max_value=len(df_display) - 1 if not df_display.empty else 0, step=1, key="edit_receber"
        )
        if not df_display.empty:
            rec = df_display.iloc[idx]
            orig_idx_candidates = df[
                (df["fornecedor"] == rec["fornecedor"]) &
                (df["valor"] == rec["valor"]) &
                (df["vencimento"] == rec["vencimento"])
            ].index
            orig_idx = orig_idx_candidates[0] if len(orig_idx_candidates) > 0 else rec.name

            colv1, colv2 = st.columns(2)
            with colv1:
                new_val = st.number_input("Valor:", value=float(rec["valor"]), key="novo_valor_receber")
                default_dt = rec["vencimento"].date() if pd.notna(rec["vencimento"]) else date.today()
                new_venc = st.date_input("Vencimento:", value=default_dt, key="novo_vencimento_receber")
            with colv2:
                estado_uni = df["estado"].dropna().astype(str).unique().tolist()
                try:
                    est_idx = estado_uni.index(str(rec["estado"]))
                except ValueError:
                    est_idx = 0
                new_estado = st.selectbox("Estado:", options=estado_uni, index=est_idx, key="novo_estado_receber")

                situ_uni = df["situacao"].dropna().astype(str).unique().tolist()
                try:
                    sit_idx = situ_uni.index(str(rec["situacao"]))
                except ValueError:
                    sit_idx = 0
                new_sit = st.selectbox("Situação:", options=situ_uni, index=sit_idx, key="nova_situacao_receber")

            if st.button("💾 Salvar Alterações", key="salvar_receber"):
                df.at[orig_idx, "valor"] = new_val
                df.at[orig_idx, "vencimento"] = pd.to_datetime(new_venc)
                df.at[orig_idx, "estado"] = new_estado
                df.at[orig_idx, "situacao"] = new_sit

                save_data(EXCEL_RECEBER, aba, df)
                df = load_data(EXCEL_RECEBER, aba)
                st.success("Registro atualizado com sucesso!")

                if view_sel == "Recebidas":
                    df_display = df[df["estado"].str.strip().str.lower() == "recebido"].copy()
                elif view_sel == "Pendentes":
                    df_display = df[df["estado"].str.strip().str.lower() != "recebido"].copy()
                else:
                    df_display = df.copy()

                if "forn" in locals() and forn != "Todos":
                    df_display = df_display[df_display["fornecedor"] == forn]
                if "status_sel" in locals() and status_sel != "Todos":
                    df_display = df_display[df_display["estado"] == status_sel]

                table_placeholder_r.dataframe(df_display[cols_para_exibir], height=250)

    st.markdown("---")

    # ======= ANEXAR DOCUMENTOS =======
    with st.expander("📎 Anexar Documentos"):
        if not df_display.empty:
            idx2 = st.number_input(
                "Índice para anexar (baseado na lista acima):",
                min_value=0, max_value=len(df_display) - 1, step=1, key="idx_anex_receber"
            )
            rec_anex = df_display.iloc[idx2]
            orig_idx_anex_candidates = df[
                (df["fornecedor"] == rec_anex["fornecedor"]) &
                (df["valor"] == rec_anex["valor"]) &
                (df["vencimento"] == rec_anex["vencimento"])
            ].index
            orig_idx_anex = orig_idx_anex_candidates[0] if len(orig_idx_anex_candidates) > 0 else rec_anex.name

            uploaded = st.file_uploader(
                "Selecione (pdf/jpg/png):", type=["pdf", "jpg", "png"], key=f"up_receber_{aba}_{idx2}"
            )
            if uploaded:
                destino = os.path.join(
                    ANEXOS_DIR, "Contas a Receber", f"Receber_{aba}_{orig_idx_anex}_{uploaded.name}"
                )
                with open(destino, "wb") as f:
                    f.write(uploaded.getbuffer())
                st.success(f"Documento salvo em: {destino}")

    st.markdown("---")

    # ======= ADICIONAR NOVA CONTA =======
    with st.expander("➕ Adicionar Nova Conta"):
        coln1, coln2 = st.columns(2)
        with coln1:
            data_nf   = st.date_input("Data N/F:", value=date.today(), key="nova_data_nf_receber")
            forma_pag = st.text_input("Descrição:", key="nova_descricao_receber")
            forn_new  = st.text_input("Fornecedor:", key="novo_fornecedor_receber")
        with coln2:
            os_new    = st.text_input("Documento/OS:", key="novo_os_receber")
            venc_new  = st.date_input("Data de Vencimento:", value=date.today(), key="novo_venc_receber")
            valor_new = st.number_input("Valor (R$):", min_value=0.0, format="%.2f", key="novo_valor_receber2")

        estado_opt  = ["A Receber", "Recebido"]
        situ_opt    = ["Em Atraso", "Recebido", "A Receber"]
        estado_new  = st.selectbox("Estado:", options=estado_opt, key="estado_novo_receber")
        situ_new    = st.selectbox("Situação:", options=situ_opt, key="situacao_novo_receber")
        boleto_file   = st.file_uploader("Boleto (opcional):",   type=["pdf", "jpg", "png"], key="boleto_novo_receber")
        comprov_file = st.file_uploader("Comprovante (opcional):", type=["pdf", "jpg", "png"], key="comprov_novo_receber")

        if st.button("➕ Adicionar Conta", key="adicionar_receber"):
            record = {
                "data_nf": data_nf,
                "forma_pagamento": forma_pag,
                "fornecedor": forn_new,
                "os": os_new,
                "vencimento": venc_new,
                "valor": valor_new,
                "estado": estado_new,
                "situacao": situ_new,
                "boleto": "",
                "comprovante": ""
            }
            if boleto_file:
                boleto_path = os.path.join(
                    ANEXOS_DIR, "Contas a Receber", f"Receber_{aba}_boleto_{boleto_file.name}"
                )
                with open(boleto_path, "wb") as fb:
                    fb.write(boleto_file.getbuffer())
                record["boleto"] = boleto_path
            if comprov_file:
                comprov_path = os.path.join(
                    ANEXOS_DIR, "Contas a Receber", f"Receber_{aba}_comprov_{comprov_file.name}"
                )
                with open(comprov_path, "wb") as fc:
                    fc.write(comprov_file.getbuffer())
                record["comprovante"] = comprov_path

            add_record(EXCEL_RECEBER, aba, record)
            st.success("Nova conta adicionada com sucesso!")

            df = load_data(EXCEL_RECEBER, aba)
            if view_sel == "Recebidas":
                df_display = df[df["estado"].str.strip().str.lower() == "recebido"].copy()
            elif view_sel == "Pendentes":
                df_display = df[df["estado"].str.strip().str.lower() != "recebido"].copy()
            else:
                df_display = df.copy()

            if "forn" in locals() and forn != "Todos":
                df_display = df_display[df_display["fornecedor"] == forn]
            if "status_sel" in locals() and status_sel != "Todos":
                df_display = df_display[df_display["estado"] == status_sel]

            table_placeholder_r.dataframe(df_display[cols_para_exibir], height=250)

    st.markdown("---")

    # ======= EXPORTAR ABA ATUAL =======
    st.subheader("💾 Exportar Aba Atual")
    try:
        df_to_save = load_data(EXCEL_RECEBER, aba)
        if not df_to_save.empty:
            save_data(EXCEL_RECEBER, aba, df_to_save)
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
