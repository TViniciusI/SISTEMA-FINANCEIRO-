import streamlit as st
import pandas as pd
import os
from datetime import datetime, date
from openpyxl import load_workbook

st.set_page_config(
    page_title="💼 Sistema Financeiro 2025",
    page_icon="💰",
    layout="wide"
)

VALID_USERS = {
    "Vinicius": "vinicius4223",
    "Flavio": "1234",
}

def check_login(username: str, password: str) -> bool:
    return VALID_USERS.get(username) == password

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = ""

if not st.session_state.logged_in:
    st.write("\n" * 5)
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

logged_user = st.session_state.username
st.sidebar.write(f"Logado como: **{logged_user}**")

EXCEL_PAGAR   = "Contas a pagar 2025.xlsx"
EXCEL_RECEBER = "Contas a receber 2025.xlsx"
ANEXOS_DIR    = "anexos"
FULL_MONTHS   = [f"{i:02d}" for i in range(1, 13)]

def get_existing_sheets(excel_path: str) -> list[str]:
    try:
        wb = pd.ExcelFile(excel_path)
        numeric_sheets = []
        for s in wb.sheet_names:
            nome = s.strip()
            if nome.lower() == "tutorial":
                continue
            if nome.isdigit():
                nome_formatado = f"{int(nome):02d}"  # transforma "4" → "04"
                numeric_sheets.append(nome_formatado)
        return sorted(set(numeric_sheets))
    except Exception:
        return []


def load_data(excel_path: str, sheet_name: str) -> pd.DataFrame:
    cols = [
        "data_nf", "forma_pagamento", "fornecedor", "os",
        "vencimento", "valor", "estado", "situacao", "boleto", "comprovante"
    ]
    if not os.path.isfile(excel_path):
        return pd.DataFrame(columns=cols + ["status_pagamento"])

    # mapeia abas numéricas ("04" → "4")
    sheet_lookup = {}
    try:
        with pd.ExcelFile(excel_path) as wb:
            for s in wb.sheet_names:
                nome = s.strip()
                if nome.lower() != "tutorial" and nome.isdigit():
                    sheet_lookup[f"{int(nome):02d}"] = nome
    except:
        return pd.DataFrame(columns=cols + ["status_pagamento"])
    if sheet_name not in sheet_lookup:
        return pd.DataFrame(columns=cols + ["status_pagamento"])
    real_sheet = sheet_lookup[sheet_name]

    try:
        df = pd.read_excel(excel_path, sheet_name=real_sheet, skiprows=7, header=0)
    except:
        return pd.DataFrame(columns=cols + ["status_pagamento"])

    # renomeia colunas
    rename_map = {}
    for col in df.columns:
        nome = str(col).strip().lower()
        if ("data" in nome and "nf" in nome) or "data da nota fiscal" in nome:
            rename_map[col] = "data_nf"
        elif "forma" in nome and "pagamento" in nome:
            rename_map[col] = "forma_pagamento"
        elif nome == "descrição":
            rename_map[col] = "forma_pagamento"
        elif nome == "fornecedor" or "cliente" in nome:
            rename_map[col] = "fornecedor"
        elif "os" in nome or nome == "documento":
            rename_map[col] = "os"
        elif "vencimento" in nome:
            rename_map[col] = "vencimento"
        elif "valor" in nome:
            rename_map[col] = "valor"
        elif nome == "estado":
            rename_map[col] = "estado"
        elif "situa" in nome:
            rename_map[col] = "situacao"
        elif "comprov" in nome:
            rename_map[col] = "comprovante"
        elif "boleto" in nome:
            rename_map[col] = "boleto"

    df = df.rename(columns=rename_map)
    df = df[[c for c in df.columns if c in cols]]

    # garante colunas mínimas
    for obrig in ["fornecedor", "valor"]:
        if obrig not in df.columns:
            df[obrig] = pd.NA

    df = df.dropna(subset=["fornecedor", "valor"], how="all").reset_index(drop=True)

    # converte tipos
    df["vencimento"] = pd.to_datetime(df["vencimento"], errors="coerce")
    df["valor"]      = pd.to_numeric(df["valor"], errors="coerce")

    # detecta modo: Pagar ou Receber
    is_receber = (excel_path == EXCEL_RECEBER)

    # monta status_pagamento
    status_list = []
    hoje = datetime.now().date()
    for _, row in df.iterrows():
        estado_atual = str(row.get("estado", "")).strip().lower()
        # já pago / recebido
        if estado_atual == ("recebido" if is_receber else "pago"):
            status_list.append("Recebido" if is_receber else "Pago")
        else:
            data_venc = row["vencimento"].date() if pd.notna(row["vencimento"]) else None
            if data_venc:
                if data_venc < hoje:
                    status_list.append("Em Atraso")
                else:
                    # pendentes: Em Aberto (pagar) ou A Receber (receber)
                    status_list.append("A Receber" if is_receber else "Em Aberto")
            else:
                status_list.append("Sem Data")

    df["status_pagamento"] = status_list
    return df


def rename_col_index(ws, target_name: str) -> int:
    for row in ws.iter_rows(min_row=1, max_row=100, min_col=1, max_col=ws.max_column):
        for cell in row:
            if cell.value and str(cell.value).strip().lower() == target_name.lower():
                return cell.column
    defaults = {"vencimento": 5, "valor": 6, "estado": 7, "situação": 8}
    return defaults.get(target_name.lower(), 1)
def save_data(excel_path: str, sheet_name: str, df: pd.DataFrame):
    """Atualiza os registros existentes sem apagar fórmulas ou layout."""
    from openpyxl import load_workbook
    from datetime import datetime, date

    wb = load_workbook(excel_path)
    ws = wb[sheet_name]

    header_row = 8
    headers = [
        str(ws.cell(row=header_row, column=col).value).strip().lower()
        for col in range(2, ws.max_column + 1)  # pula coluna 1
    ]

    field_map = {
        "data_nf": [
            "data documento", "data_nf", "data n/f", "data n/fornecedor",
            "data da nota fiscal", "data n/ffornecedor"
        ],
        "forma_pagamento": ["descrição", "forma_pagamento", "forma de pagamento"],
        "fornecedor": ["fornecedor"],
        "os": ["documento", "os", "os interna", "os_interna"],
        "vencimento": ["vencimento"],
        "valor": ["valor"],
        "estado": ["estado"],
        "boleto": ["boleto", "boleto anexo"],
        "comprovante": ["comprovante", "comprovante de pagto"]
    }

    col_pos = {}
    for key, names in field_map.items():
        idx = next((i for i, h in enumerate(headers) if h in names), None)
        col_pos[key] = idx + 2 if idx is not None else None  # +2 porque headers pula col 1

    for i, row in df.iterrows():
        excel_row = header_row + 1 + i
        for key, col in col_pos.items():
            if not col or key == "situacao":
                continue
            val = row.get(key, "")

            if key in ("data_nf", "vencimento"):
                try:
                    val = pd.to_datetime(val, errors="coerce")
                    if pd.notna(val):
                        val = val.to_pydatetime()
                    else:
                        continue
                except:
                    continue

            if key == "valor":
                try:
                    val = float(val)
                except:
                    val = None

            ws.cell(row=excel_row, column=col, value=val)

    wb.save(excel_path)

def add_record(excel_path: str, sheet_name: str, record: dict):
    from openpyxl import load_workbook
    import pandas as pd
    from datetime import datetime

    wb = load_workbook(excel_path)
    if sheet_name not in wb.sheetnames:
        # copia primeira aba numérica como template
        numeric = [s for s in wb.sheetnames if s.isdigit()]
        template_ws = wb[numeric[0]] if numeric else wb[wb.sheetnames[0]]
        ws = wb.copy_worksheet(template_ws)
        ws.title = sheet_name
    else:
        ws = wb[sheet_name]

    header_row = 8
    # lê títulos para encontrar colunas
    headers = [
        str(ws.cell(row=header_row, column=col).value).strip().lower()
        for col in range(2, ws.max_column + 1)
    ]

    field_map = {
        "data_nf":       ["data documento", "data_nf", "data n/f", "data da nota fiscal"],
        "forma_pagamento": ["descrição", "forma_pagamento", "forma de pagamento"],
        "fornecedor":    ["fornecedor"],
        "os":            ["documento", "os", "os interna"],
        "vencimento":    ["vencimento"],
        "valor":         ["valor"],
        "estado":        ["estado"],
        "boleto":        ["boleto", "boleto anexo"],
        "comprovante":   ["comprovante", "comprovante de pagto"]
    }

    # mapeia posição de cada campo
    col_pos = {}
    for key, names in field_map.items():
        idx = next((i for i, h in enumerate(headers) if h in names), None)
        col_pos[key] = idx + 2 if idx is not None else None

    # fallback: garante que col_forn seja pelo menos 2
    col_forn = col_pos.get("fornecedor")
    if col_forn is None or col_forn < 2:
        col_forn = 2

    # encontra primeira linha vazia a partir de header_row+1
    next_row = ws.max_row + 1
    for r in range(header_row + 1, ws.max_row + 2):
        if not ws.cell(row=r, column=col_forn).value:
            next_row = r
            break

    # preenche os valores
    for key, col in col_pos.items():
        if not col or key == "situacao":
            continue
        val = record.get(key, "")
        if key in ("data_nf", "vencimento"):
            try:
                dt = pd.to_datetime(val, errors="coerce")
                if pd.notna(dt):
                    val = dt.to_pydatetime()
                else:
                    continue
            except:
                continue
        if key == "valor":
            try:
                val = float(val)
            except:
                val = None
        ws.cell(row=next_row, column=col, value=val)

    wb.save(excel_path)


# garante pastas de anexos
for pasta in ["Contas a Pagar", "Contas a Receber"]:
    os.makedirs(os.path.join(ANEXOS_DIR, pasta), exist_ok=True)
st.sidebar.markdown(
    """
    ## 📂 Navegação  
    Selecione a seção desejada para visualizar e gerenciar  
    suas contas a pagar e receber.  
    """,
    unsafe_allow_html=True
)

page = st.sidebar.radio("", ["Dashboard", "Contas a Pagar", "Contas a Receber"], index=0)

st.markdown("""
<div style="text-align: center; color: #4B8BBE; margin-bottom: 10px;">
    <h1>💼 Sistema Financeiro 2025</h1>
    <p style="color: #555; font-size: 16px;">Dashboard avançado com estatísticas e gráficos interativos.</p>
</div>
""", unsafe_allow_html=True)
st.markdown("---")

if page == "Dashboard":
    st.subheader("📊 Painel de Controle Financeiro Avançado")
    if not os.path.isfile(EXCEL_PAGAR):
        st.error(f"Arquivo '{EXCEL_PAGAR}' não encontrado. Verifique o caminho.")
        st.stop()
    if not os.path.isfile(EXCEL_RECEBER):
        st.error(f"Arquivo '{EXCEL_RECEBER}' não encontrado. Verifique o caminho.")
        st.stop()
    sheets_p = get_existing_sheets(EXCEL_PAGAR)
    sheets_r = get_existing_sheets(EXCEL_RECEBER)
    tabs = st.tabs(["📥 Contas a Pagar", "📤 Contas a Receber"])
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
            # Substitua daqui ↓
            monthly_group_r = (
                df_all_r
                .groupby("mes_ano")
                .agg(
                    total_mes      = ("valor", "sum"),
                    recebidos_mes  = ("valor", lambda x: x[df_all_r.loc[x.index, "status_pagamento"] == "Recebido"].sum()),
                    pendentes_mes  = ("valor", lambda x: x[df_all_r.loc[x.index, "status_pagamento"] != "Recebido"].sum())
                )
                .reset_index()
            )
            # Até aqui ↑
        
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

elif page == "Contas a Pagar":
    st.subheader("🗂️ Contas a Pagar")
    if not os.path.isfile(EXCEL_PAGAR):
        st.error(f"Arquivo '{EXCEL_PAGAR}' não encontrado. Verifique o caminho.")
        st.stop()
    existing = get_existing_sheets(EXCEL_PAGAR)
    aba = st.selectbox("Selecione o mês:", FULL_MONTHS, index=0)
    df = load_data(EXCEL_PAGAR, aba)
    if df.empty:
        st.info("Nenhum registro encontrado para este mês (ou a aba não existia).")
    view_sel = st.radio("Visualizar:", ["Todos", "Pagas", "Pendentes"], horizontal=True)
    if view_sel == "Pagas":
        df_display = df[df["estado"].astype(str).str.strip().str.lower() == "pago"].copy()
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


    with st.expander("✏️ Editar Registro"):
        idx = st.number_input(
            "Índice da linha (baseado na lista acima):",
            min_value=0,
            max_value=len(df_display) - 1 if not df_display.empty else 0,
            step=1,
            key="edit_pagar"
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
                new_val = st.number_input(
                    "Valor:",
                    value=float(rec["valor"]),
                    key="novo_valor_pagar"
                )
                default_dt = (
                    rec["vencimento"].date()
                    if pd.notna(rec["vencimento"])
                    else date.today()
                )
                new_venc = st.date_input(
                    "Vencimento:",
                    value=default_dt,
                    key="novo_vencimento_pagar"
                )
            with colv2:
                # lista fixa de estados
                estado_opt = ["Em Aberto", "Pago"]
                try:
                    default_idx = estado_opt.index(str(rec["estado"]))
                except ValueError:
                    default_idx = 0
                new_estado = st.selectbox(
                    "Estado:",
                    options=estado_opt,
                    index=default_idx,
                    key="novo_estado_pagar"
                )

                # lista fixa de situações
                situ_opt = ["Em Atraso", "Pago", "Em Aberto"]
                try:
                    sit_idx = situ_opt.index(str(rec["situacao"]))
                except ValueError:
                    sit_idx = 0
                new_sit = st.selectbox(
                    "Situação:",
                    options=situ_opt,
                    index=sit_idx,
                    key="nova_situacao_pagar"
                )

            if st.button("💾 Salvar Alterações", key="salvar_pagar"):
                df.at[orig_idx, "valor"] = new_val
                df.at[orig_idx, "vencimento"] = pd.to_datetime(new_venc)
                df.at[orig_idx, "estado"] = new_estado
                df.at[orig_idx, "situacao"] = new_sit
                save_data(EXCEL_PAGAR, aba, df)

                df = load_data(EXCEL_PAGAR, aba)
                st.success("Registro atualizado com sucesso!")

                # reaplica view_sel e filtros
                if view_sel == "Pagas":
                    df_display = df[df["status_pagamento"] == "Pago"].copy()
                elif view_sel == "Pendentes":
                    df_display = df[df["status_pagamento"] != "Pago"].copy()
                else:
                    df_display = df.copy()
                if forn != "Todos":
                    df_display = df_display[df_display["fornecedor"] == forn]
                if status_sel != "Todos":
                    df_display = df_display[df_display["status_pagamento"] == status_sel]

                table_placeholder.dataframe(
                    df_display[cols_para_exibir],
                    height=250
                )
    # 🗑️ Remover Registro (Contas a Pagar)
    with st.expander("🗑️ Remover Registro"):
        if not df_display.empty:
            idx_rem = st.number_input(
                "Índice da linha para remover:",
                min_value=0,
                max_value=len(df_display) - 1,
                step=1,
                key="remover_pagar"
            )
            if st.button("Remover", key="btn_remover_pagar"):
                # 1) Remove do DataFrame em memória
                rec_rem = df_display.iloc[idx_rem]
                orig_idx = rec_rem.name
                df_display = df_display.drop(orig_idx).reset_index(drop=True)

                # 2) Reescreve a aba no Excel usando pandas
                import pandas as pd
                with pd.ExcelWriter(EXCEL_PAGAR, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                    df_display.to_excel(
                        writer,
                        sheet_name=aba,
                        index=False,
                        startrow=7  # mantém cabeçalho das 8 primeiras linhas
                    )

                st.success("Registro removido com sucesso!")

                # 3) Recarrega e exibe a tabela
                df = load_data(EXCEL_PAGAR, aba)
                if view_sel == "Pagas":
                    df_display = df[df["status_pagamento"] == "Pago"].copy()
                elif view_sel == "Pendentes":
                    df_display = df[df["status_pagamento"] != "Pago"].copy()
                else:
                    df_display = df.copy()
                if forn != "Todos":
                    df_display = df_display[df_display["fornecedor"] == forn]

                cols_show       = ["data_nf","fornecedor","valor","vencimento","estado","status_pagamento"]
                cols_to_display = [c for c in cols_show if c in df_display.columns]
                table_placeholder.dataframe(df_display[cols_to_display], height=250)


    st.markdown("---")
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
    # ➕ Adicionar Nova Conta (Contas a Pagar)
    with st.expander("➕ Adicionar Nova Conta"):
        coln1, coln2 = st.columns(2)
        with coln1:
            data_nf   = st.date_input(
                "Data N/F:",
                value=date.today(),
                key="nova_data_nf_pagar"
            )
            forma_pag = st.text_input(
                "Descrição:",
                key="nova_descricao_pagar"
            )
            forn_new  = st.text_input(
                "Fornecedor:",
                key="novo_fornecedor_pagar"
            )
        with coln2:
            os_new    = st.text_input(
                "Documento/OS:",
                key="novo_os_pagar"
            )
            venc_new  = st.date_input(
                "Data de Vencimento:",
                value=date.today(),
                key="novo_venc_pagar"
            )
            valor_new = st.number_input(
                "Valor (R$):",
                min_value=0.0,
                format="%.2f",
                key="novo_valor_pagar2"
            )

        # Opções para Contas a Pagar
        estado_opt = ["Em Aberto", "Pago"]
        situ_opt   = ["Em Atraso", "Pago", "Em Aberto"]
        estado_new = st.selectbox(
            "Estado:",
            options=estado_opt,
            key="estado_novo_pagar"
        )
        situ_new   = st.selectbox(
            "Situação:",
            options=situ_opt,
            key="situacao_novo_pagar"
        )

        boleto_file   = st.file_uploader(
            "Boleto (opcional):",
            type=["pdf", "jpg", "png"],
            key="boleto_novo_pagar"
        )
        comprov_file = st.file_uploader(
            "Comprovante (opcional):",
            type=["pdf", "jpg", "png"],
            key="comprov_novo_pagar"
        )

        if st.button("➕ Adicionar Conta", key="adicionar_pagar"):
            record = {
                "data_nf":        data_nf,
                "forma_pagamento": forma_pag,
                "fornecedor":     forn_new,
                "os":             os_new,
                "vencimento":     venc_new,
                "valor":          valor_new,
                "estado":         estado_new,
                "situacao":       situ_new,
                "boleto":         "",
                "comprovante":    ""
            }
            # salva anexos em pasta correta
            if boleto_file:
                boleto_path = os.path.join(
                    ANEXOS_DIR, "Contas a Pagar",
                    f"Pagar_{aba}_boleto_{boleto_file.name}"
                )
                with open(boleto_path, "wb") as fb:
                    fb.write(boleto_file.getbuffer())
                record["boleto"] = boleto_path
            if comprov_file:
                comprov_path = os.path.join(
                    ANEXOS_DIR, "Contas a Pagar",
                    f"Pagar_{aba}_comprov_{comprov_file.name}"
                )
                with open(comprov_path, "wb") as fc:
                    fc.write(comprov_file.getbuffer())
                record["comprovante"] = comprov_path

            add_record(EXCEL_PAGAR, aba, record)
            st.success("Nova conta adicionada com sucesso!")

            # botões de download dos anexos
            if record.get("boleto"):
                with open(record["boleto"], "rb") as f:
                    st.download_button(
                        label="📥 Baixar Boleto",
                        data=f.read(),
                        file_name=os.path.basename(record["boleto"]),
                        mime="application/octet-stream",
                        key=f"dl_boleto_pagar_{aba}"
                    )
            if record.get("comprovante"):
                with open(record["comprovante"], "rb") as f:
                    st.download_button(
                        label="📥 Baixar Comprovante",
                        data=f.read(),
                        file_name=os.path.basename(record["comprovante"]),
                        mime="application/octet-stream",
                        key=f"dl_comprov_pagar_{aba}"
                    )

            # Recarrega tabela
            df = load_data(EXCEL_PAGAR, aba)
            cols_show       = ["data_nf","fornecedor","valor","vencimento","estado","status_pagamento"]
            cols_to_display = [c for c in cols_show if c in df.columns]
            table_placeholder.dataframe(df[cols_to_display], height=250)


    st.markdown("---")
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
    st.markdown("---")


elif page == "Contas a Receber":
    st.subheader("🗂️ Contas a Receber")

    # 1) Seletor de mês e carregamento dos dados
    aba = st.selectbox(
        "Selecione o mês:",
        FULL_MONTHS,
        index=FULL_MONTHS.index(date.today().strftime("%m"))
    )
    df = load_data(EXCEL_RECEBER, aba)

    # 2) Definição de df_display via view_sel
    view_sel = st.radio(
        "Visualizar:",
        ["Todos", "Recebidas", "Pendentes"],
        horizontal=True
    )
    if view_sel == "Recebidas":
        df_display = df[df["status_pagamento"] == "Recebido"].copy()
    elif view_sel == "Pendentes":
        df_display = df[df["status_pagamento"] != "Recebido"].copy()
    else:
        df_display = df.copy()

    # 3) Filtros adicionais
    with st.expander("🔍 Filtros"):
        col1, col2 = st.columns(2)
        with col1:
            forn = st.selectbox(
                "Fornecedor",
                ["Todos"] + sorted(df["fornecedor"].dropna().astype(str).unique())
            )
        with col2:
            status_sel = st.selectbox(
                "Status",
                ["Todos"] + sorted(df["status_pagamento"].dropna().unique())
            )

    if forn != "Todos":
        df_display = df_display[df_display["fornecedor"] == forn]
    if status_sel != "Todos":
        df_display = df_display[df_display["status_pagamento"] == status_sel]

    st.markdown("<hr>", unsafe_allow_html=True)
    if df_display.empty:
        st.warning("Nenhum registro para os filtros selecionados.")
    else:
        cols_show = ["data_nf", "fornecedor", "valor", "vencimento", "estado", "status_pagamento"]
        # filtra apenas as colunas presentes no DataFrame
        cols_to_display = [c for c in cols_show if c in df_display.columns]
        table_placeholder_r = st.empty()
        table_placeholder_r.dataframe(df_display[cols_to_display], height=250)
    st.markdown("---")

    
    # ✏️ Editar Registro
    with st.expander("✏️ Editar Registro"):
        if df_display.empty:
            st.info("Nenhum registro para editar.")
        else:
            # Seleção de índice na tabela filtrada
            idx = st.number_input(
                "Índice da linha (baseado na lista acima):",
                min_value=0,
                max_value=len(df_display) - 1,
                step=1,
                key="edit_receber"
            )

            # Recupera a linha e seu índice original
            rec = df_display.iloc[idx]
            orig_idx = rec.name

            # Campos para editar valor e vencimento
            col1, col2 = st.columns(2)
            with col1:
                new_val = st.number_input(
                    "Valor:",
                    value=float(rec["valor"]),
                    key="novo_valor_receber"
                )
                new_venc = st.date_input(
                    "Vencimento:",
                    rec["vencimento"].date() if pd.notna(rec["vencimento"]) else date.today(),
                    key="novo_vencimento_receber"
                )
            with col2:
                estado_opt = ["A Receber", "Recebido"]
                situ_opt   = ["Em Atraso", "Recebido", "A Receber"]
                new_estado = st.selectbox(
                    "Estado:",
                    options=estado_opt,
                    index=estado_opt.index(rec["estado"]) if rec["estado"] in estado_opt else 0,
                    key="novo_estado_receber"
                )
                new_sit = st.selectbox(
                    "Situação:",
                    options=situ_opt,
                    index=situ_opt.index(rec["situacao"]) if rec["situacao"] in situ_opt else 0,
                    key="nova_situacao_receber"
                )

            # Botão de salvar
            if st.button("💾 Salvar Alterações", key="salvar_receber"):
                # Atualiza o DataFrame original
                df.at[orig_idx, "valor"]      = new_val
                df.at[orig_idx, "vencimento"] = pd.to_datetime(new_venc)
                df.at[orig_idx, "estado"]     = new_estado
                df.at[orig_idx, "situacao"]   = new_sit
                save_data(EXCEL_RECEBER, aba, df)
                st.success("Registro atualizado com sucesso!")

                # Recarrega e reaplica filtros
                df = load_data(EXCEL_RECEBER, aba)
                if view_sel == "Recebidas":
                    df_display = df[df["status_pagamento"] == "Recebido"].copy()
                elif view_sel == "Pendentes":
                    df_display = df[df["status_pagamento"] != "Recebido"].copy()
                else:
                    df_display = df.copy()
                if forn != "Todos":
                    df_display = df_display[df_display["fornecedor"] == forn]
                if status_sel != "Todos":
                    df_display = df_display[df_display["status_pagamento"] == status_sel]

                cols_show = ["data_nf", "fornecedor", "valor", "vencimento", "estado", "status_pagamento"]
                cols_to_display = [c for c in cols_show if c in df_display.columns]
                table_placeholder_r.dataframe(df_display[cols_to_display], height=250)

 # 🗑️ Remover Registro
    with st.expander("🗑️ Remover Registro"):
        if not df_display.empty:
            idx_r = st.number_input(
                "Índice para remover:",
                min_value=0,
                max_value=len(df_display) - 1,
                step=1,
                key="remover_receber"
            )
            if st.button("Remover", key="btn_remover_receber"):
                rec = df_display.iloc[idx_r]
                orig_idx = df.index[
                    (df["fornecedor"] == rec["fornecedor"]) &
                    (df["valor"] == rec["valor"]) &
                    (df["vencimento"] == rec["vencimento"])
                ][0]
                wb = load_workbook(EXCEL_RECEBER)
                ws = wb[aba]
                ws.delete_rows(8 + 1 + orig_idx)
                wb.save(EXCEL_RECEBER)
                st.success("Registro removido com sucesso!")

                df = load_data(EXCEL_RECEBER, aba)
                # reaplica filtros
                if view_sel == "Recebidas":
                    df_display = df[df["status_pagamento"] == "Recebido"].copy()
                elif view_sel == "Pendentes":
                    df_display = df[df["status_pagamento"] != "Recebido"].copy()
                else:
                    df_display = df.copy()
                if forn != "Todos":
                    df_display = df_display[df_display["fornecedor"] == forn]
                if status_sel != "Todos":
                    df_display = df_display[df_display["status_pagamento"] == status_sel]
                table_placeholder_r.dataframe(df_display[cols_show], height=250)
       # 📎 Anexar Documentos
    with st.expander("📎 Anexar Documentos"):
        if not df_display.empty:
            idx2 = st.number_input(
                "Índice para anexar:",
                min_value=0,
                max_value=len(df_display) - 1,
                step=1,
                key=f"idx_anex_receber_{aba}"
            )
            rec2 = df_display.iloc[idx2]
            orig2 = rec2.name  # usa o índice da linha filtrada

            up = st.file_uploader(
                "Selecione (pdf/jpg/png):",
                type=["pdf", "jpg", "png"],
                key=f"up_receber_{aba}_{idx2}"
            )
            if up:
                destino = os.path.join(
                    ANEXOS_DIR,
                    "Contas a Receber",
                    f"Receber_{aba}_{orig2}_{up.name}"
                )
                with open(destino, "wb") as f:
                    f.write(up.getbuffer())
                st.success(f"Documento salvo em: {destino}")

    # ➕ Adicionar Nova Conta
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

        estado_opt = ["A Receber", "Recebido"]
        situ_opt   = ["Em Atraso", "Recebido", "A Receber"]
        estado_new = st.selectbox("Estado:", options=estado_opt, key="estado_novo_receber")
        situ_new   = st.selectbox("Situação:", options=situ_opt,   key="situacao_novo_receber")
        boleto_file   = st.file_uploader("Boleto (opcional):",   type=["pdf","jpg","png"], key="boleto_novo_receber")
        comprov_file = st.file_uploader("Comprovante (opcional):", type=["pdf","jpg","png"], key="comprov_novo_receber")

        if st.button("➕ Adicionar Conta", key="adicionar_receber"):
            record = {
                "data_nf":        data_nf,
                "forma_pagamento": forma_pag,
                "fornecedor":     forn_new,
                "os":             os_new,
                "vencimento":     venc_new,
                "valor":          valor_new,
                "estado":         estado_new,
                "situacao":       situ_new,
                "boleto":         "",
                "comprovante":    ""
            }
            if boleto_file:
                p = os.path.join(
                    ANEXOS_DIR,
                    "Contas a Receber",
                    f"Receber_{aba}_boleto_{boleto_file.name}"
                )
                with open(p, "wb") as fb:
                    fb.write(boleto_file.getbuffer())
                record["boleto"] = p
            if comprov_file:
                p = os.path.join(
                    ANEXOS_DIR,
                    "Contas a Receber",
                    f"Receber_{aba}_comprov_{comprov_file.name}"
                )
                with open(p, "wb") as fc:
                    fc.write(comprov_file.getbuffer())
                record["comprovante"] = p

            add_record(EXCEL_RECEBER, aba, record)
            st.success("Nova conta adicionada com sucesso!")

            if record.get("boleto"):
                with open(record["boleto"], "rb") as f:
                    st.download_button(
                        label="📥 Baixar Boleto",
                        data=f.read(),
                        file_name=os.path.basename(record["boleto"]),
                        mime="application/octet-stream",
                        key=f"dl_boleto_{aba}"
                    )
            if record.get("comprovante"):
                with open(record["comprovante"], "rb") as f:
                    st.download_button(
                        label="📥 Baixar Comprovante",
                        data=f.read(),
                        file_name=os.path.basename(record["comprovante"]),
                        mime="application/octet-stream",
                        key=f"dl_comprov_{aba}"
                    )

            # Recarrega e exibe o DataFrame atualizado
            df = load_data(EXCEL_RECEBER, aba)
            # define quais colunas mostrar, verificando quais existem de fato
            cols_show      = ["data_nf", "fornecedor", "valor", "vencimento", "estado", "status_pagamento"]
            cols_to_display = [c for c in cols_show if c in df.columns]
            table_placeholder_r.dataframe(df[cols_to_display], height=250)

    st.markdown("---")
    st.subheader("💾 Exportar Aba Atual")
    try:
        df_to_save = load_data(EXCEL_RECEBER, aba)
        if not df_to_save.empty:
            save_data(EXCEL_RECEBER, aba, df_to_save)
        with open(EXCEL_RECEBER, "rb") as fx:
            st.download_button(
                label=f"Exportar '{aba}'",
                data=fx.read(),
                file_name=f"Contas a Receber - {aba}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Erro ao preparar download: {e}")

st.markdown("""
<div style="text-align: center; font-size:12px; color:gray; margin-top: 20px;">
    <p>© 2025 Desenvolvido por Vinicius Magalhães</p>
</div>
""", unsafe_allow_html=True)
