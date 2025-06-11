import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
import os
from openpyxl import load_workbook

# Configuração da página
st.set_page_config(
    page_title="💼 Sistema Financeiro 2025",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
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

# Constantes no início do arquivo (após as imports)
EXCEL_PAGAR = "Contas a pagar 2025.xlsx"
EXCEL_RECEBER = "Contas a receber 2025.xlsx"
ANEXOS_DIR = "anexos"  # Esta linha estava faltando
FULL_MONTHS = [f"{i:02d}" for i in range(1, 13)]

# Garante pastas de anexos (esta parte deve vir DEPOIS de definir ANEXOS_DIR)
for pasta in ["Contas a Pagar", "Contas a Receber"]:
    os.makedirs(os.path.join(ANEXOS_DIR, pasta), exist_ok=True)



st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    /* ------------------------- Cabeçalho ------------------------- */
    .header {
        background: linear-gradient(135deg, #3c3b92, #6051db);
        color: white;
        padding: 1.5rem;
        border-radius: 14px;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.25);
        margin-bottom: 2rem;
    }

    /* ------------------------- Cards de Métricas ------------------------- */
    .metric-card {
        border-radius: 14px;
        padding: 1.5rem;
        border-left: 4px solid #6051db;
        transition: transform 0.25s ease;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.05);
    }

    .metric-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 8px 18px rgba(0, 0, 0, 0.1);
    }

    .metric-value {
        font-size: 2.4rem;
        font-weight: 600;
        margin-bottom: 0.25rem;
    }

    .metric-label {
        font-size: 0.85rem;
        text-transform: uppercase;
        letter-spacing: 0.7px;
    }

    /* Tema Claro */
    .stApp:not([data-theme="dark"]) .metric-card {
        background: #f8f9fc;
    }

    .stApp:not([data-theme="dark"]) .metric-value {
        color: #2c3e50;
    }

    .stApp:not([data-theme="dark"]) .metric-label {
        color: #7b8a97;
    }

    /* Tema Escuro */
    .stApp[data-theme="dark"] .metric-card {
        background: #20212b;
    }

    .stApp[data-theme="dark"]) .metric-value {
        color: #f0f2f5;
    }

    .stApp[data-theme="dark"] .metric-label {
        color: #b0b3bd;
    }

    /* ------------------------- Tabs ------------------------- */
    .stTabs [role="tablist"] {
        gap: 10px;
    }

    .stTabs [role="tab"] {
        padding: 10px 20px;
        border-radius: 10px 10px 0 0;
        font-weight: 500;
        transition: all 0.2s ease-in-out;
        border: none;
    }

    .stApp:not([data-theme="dark"]) .stTabs [role="tab"] {
        background: #e0e3ec;
        color: #2c3e50;
    }

    .stApp[data-theme="dark"] .stTabs [role="tab"] {
        background: #2a2b38;
        color: #ccc;
    }

    .stTabs [role="tab"][aria-selected="true"] {
        background: #6051db !important;
        color: #fff !important;
    }

    /* ------------------------- Gráfico Container ------------------------- */
    .chart-container {
        border-radius: 14px;
        padding: 1.5rem;
        margin-bottom: 2rem;
    }

    .stApp:not([data-theme="dark"]) .chart-container {
        background: #ffffff;
        box-shadow: 0 1px 6px rgba(0, 0, 0, 0.04);
    }

    .stApp[data-theme="dark"] .chart-container {
        background: #1d1e26;
        box-shadow: 0 1px 10px rgba(0, 0, 0, 0.2);
    }

    /* ------------------------- Botões ------------------------- */
    .stButton > button, .stDownloadButton > button {
        border-radius: 8px;
        padding: 10px 20px;
        font-weight: 500;
        transition: all 0.25s ease;
        border: none;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }

    .stApp:not([data-theme="dark"]) .stButton > button,
    .stApp:not([data-theme="dark"]) .stDownloadButton > button {
        background-color: #4e54c8;
        color: white;
    }

    .stApp[data-theme="dark"] .stButton > button,
    .stApp[data-theme="dark"] .stDownloadButton > button {
        background-color: #7a7ef7;
        color: white;
    }

    .stButton > button:hover,
    .stDownloadButton > button:hover {
        filter: brightness(0.9);
        transform: scale(1.02);
    }

</style>
""", unsafe_allow_html=True)

def get_existing_sheets(excel_path: str) -> list[str]:
    try:
        wb = pd.ExcelFile(excel_path)
        numeric_sheets = []
        for s in wb.sheet_names:
            nome = s.strip()
            if nome.lower() == "tutorial":
                continue
            if nome.isdigit():
                nome_formatado = f"{int(nome):02d}"
                numeric_sheets.append(nome_formatado)
        return sorted(set(numeric_sheets))
    except Exception as e:
        st.error(f"Erro ao ler abas do arquivo: {e}")
        return []

def load_data(excel_path: str, sheet_name: str) -> pd.DataFrame:
    cols = [
        "data_nf", "forma_pagamento", "fornecedor", "os",
        "vencimento", "valor", "estado", "situacao", "boleto", "comprovante"
    ]
    
    if not os.path.isfile(excel_path):
        return pd.DataFrame(columns=cols + ["status_pagamento"])

    try:
        # Mapeia abas numéricas ("04" → "4")
        sheet_lookup = {}
        with pd.ExcelFile(excel_path) as wb:
            for s in wb.sheet_names:
                nome = s.strip()
                if nome.lower() != "tutorial" and nome.isdigit():
                    sheet_lookup[f"{int(nome):02d}"] = nome
        
        if sheet_name not in sheet_lookup:
            return pd.DataFrame(columns=cols + ["status_pagamento"])
            
        real_sheet = sheet_lookup[sheet_name]
        df = pd.read_excel(excel_path, sheet_name=real_sheet, skiprows=7, header=0)

        # Renomeia colunas
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

        # Garante colunas mínimas
        for obrig in ["fornecedor", "valor"]:
            if obrig not in df.columns:
                df[obrig] = pd.NA

        df = df.dropna(subset=["fornecedor", "valor"], how="all").reset_index(drop=True)

        # Converte tipos
        df["vencimento"] = pd.to_datetime(df["vencimento"], errors="coerce")
        df["valor"] = pd.to_numeric(df["valor"], errors="coerce")

        # Detecta modo: Pagar ou Receber
        is_receber = (excel_path == EXCEL_RECEBER)

        # Monta status_pagamento
        status_list = []
        hoje = datetime.now().date()
        for _, row in df.iterrows():
            estado_atual = str(row.get("estado", "")).strip().lower()
            if estado_atual == ("recebido" if is_receber else "pago"):
                status_list.append("Recebido" if is_receber else "Pago")
            else:
                data_venc = row["vencimento"].date() if pd.notna(row["vencimento"]) else None
                if data_venc:
                    if data_venc < hoje:
                        status_list.append("Em Atraso")
                    else:
                        status_list.append("A Receber" if is_receber else "Em Aberto")
                else:
                    status_list.append("Sem Data")

        df["status_pagamento"] = status_list
        return df

    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return pd.DataFrame(columns=cols + ["status_pagamento"])

def save_data(excel_path: str, sheet_name: str, df: pd.DataFrame) -> bool:
    try:
        wb = load_workbook(excel_path)
        
        if sheet_name not in wb.sheetnames:
            st.error(f"A aba '{sheet_name}' não existe no arquivo.")
            return False
            
        ws = wb[sheet_name]
        header_row = 8
        
        headers = [
            str(ws.cell(row=header_row, column=col).value).strip().lower()
            for col in range(2, ws.max_column + 1)
        ]

        field_map = {
            "data_nf": ["data documento", "data_nf", "data n/f", "data da nota fiscal"],
            "forma_pagamento": ["descrição", "forma_pagamento", "forma de pagamento"],
            "fornecedor": ["fornecedor"],
            "os": ["documento", "os", "os interna"],
            "vencimento": ["vencimento"],
            "valor": ["valor"],
            "estado": ["estado"],
            "boleto": ["boleto", "boleto anexo"],
            "comprovante": ["comprovante", "comprovante de pagto"]
        }

        col_pos = {}
        for key, names in field_map.items():
            idx = next((i for i, h in enumerate(headers) if h in names), None)
            col_pos[key] = idx + 2 if idx is not None else None

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
                elif key == "valor":
                    try:
                        val = float(val)
                    except:
                        val = None

                ws.cell(row=excel_row, column=col, value=val)

        wb.save(excel_path)
        return True
        
    except Exception as e:
        st.error(f"Erro ao salvar dados: {e}")
        return False

def add_record(excel_path: str, sheet_name: str, record: dict) -> bool:
    try:
        wb = load_workbook(excel_path)
        
        if sheet_name not in wb.sheetnames:
            numeric = [s for s in wb.sheetnames if s.isdigit()]
            template_ws = wb[numeric[0]] if numeric else wb[wb.sheetnames[0]]
            ws = wb.copy_worksheet(template_ws)
            ws.title = sheet_name
        else:
            ws = wb[sheet_name]

        header_row = 8
        headers = [
            str(ws.cell(row=header_row, column=col).value).strip().lower()
            for col in range(2, ws.max_column + 1)
        ]

        field_map = {
            "data_nf": ["data documento", "data_nf", "data n/f", "data da nota fiscal"],
            "forma_pagamento": ["descrição", "forma_pagamento", "forma de pagamento"],
            "fornecedor": ["fornecedor"],
            "os": ["documento", "os", "os interna"],
            "vencimento": ["vencimento"],
            "valor": ["valor"],
            "estado": ["estado"],
            "boleto": ["boleto", "boleto anexo"],
            "comprovante": ["comprovante", "comprovante de pagto"]
        }

        col_pos = {}
        for key, names in field_map.items():
            idx = next((i for i, h in enumerate(headers) if h in names), None)
            col_pos[key] = idx + 2 if idx is not None else None

        col_forn = col_pos.get("fornecedor", 2)
        
        next_row = ws.max_row + 1
        for r in range(header_row + 1, ws.max_row + 2):
            if not ws.cell(row=r, column=col_forn).value:
                next_row = r
                break

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
            elif key == "valor":
                try:
                    val = float(val)
                except:
                    val = None

            ws.cell(row=next_row, column=col, value=val)

        wb.save(excel_path)
        return True
        
    except Exception as e:
        st.error(f"Erro ao adicionar registro: {e}")
        return False

# Garante pastas de anexos
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

# Dashboard Modernizado
if page == "Dashboard":
    # Cabeçalho moderno
    st.markdown("""
    <div class="header">
        <h1 style="color: white; margin: 0;">💼 Sistema Financeiro 2025</h1>
        <p style="color: rgba(255, 255, 255, 0.8); margin: 0.5rem 0 0;">Dashboard avançado com estatísticas e gráficos interativos</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Verificação dos arquivos
    if not os.path.isfile(EXCEL_PAGAR):
        st.error(f"Arquivo '{EXCEL_PAGAR}' não encontrado. Verifique o caminho.")
    if not os.path.isfile(EXCEL_RECEBER):
        st.error(f"Arquivo '{EXCEL_RECEBER}' não encontrado. Verifique o caminho.")
    if not os.path.isfile(EXCEL_PAGAR) or not os.path.isfile(EXCEL_RECEBER):
        st.stop()
    
    sheets_p = get_existing_sheets(EXCEL_PAGAR)
    sheets_r = get_existing_sheets(EXCEL_RECEBER)
    
    # Layout com tabs modernas
    tab1, tab2 = st.tabs(["📥 Contas a Pagar", "📤 Contas a Receber"])
    
    with tab1:
        if not sheets_p:
            st.warning("Nenhuma aba válida encontrada em Contas a Pagar")
        else:
            df_all_p = pd.concat([load_data(EXCEL_PAGAR, s) for s in sheets_p], ignore_index=True)
            
            if df_all_p.empty:
                st.info("Nenhum dado encontrado nas planilhas de Contas a Pagar")
            else:
                # Métricas principais
                total_p = df_all_p["valor"].sum()
                num_lanc_p = len(df_all_p)
                media_p = df_all_p["valor"].mean() if num_lanc_p else 0
                atrasados_p = df_all_p[df_all_p["status_pagamento"] == "Em Atraso"]
                num_atras_p = len(atrasados_p)
                perc_atras_p = (num_atras_p / num_lanc_p * 100) if num_lanc_p else 0
                
                # Layout de métricas
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">Total a Pagar</div>
                        <div class="metric-value">R$ {total_p:,.2f}</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">Lançamentos</div>
                        <div class="metric-value">{num_lanc_p}</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col3:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">Média por Conta</div>
                        <div class="metric-value">R$ {media_p:,.2f}</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col4:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">Em Atraso</div>
                        <div class="metric-value">{perc_atras_p:.1f}%</div>
                        <div style="font-size: 0.8rem; color: {'#e74c3c' if perc_atras_p > 10 else '#27ae60'}">
                            ({num_atras_p} conta{'s' if num_atras_p != 1 else ''})
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                
                st.markdown("---")
                
                # Gráfico de distribuição por status
                st.markdown("#### 📊 Distribuição por Status")
                status_counts_p = (
                    df_all_p["status_pagamento"]
                    .value_counts()
                    .rename_axis("status")
                    .reset_index(name="contagem")
                )
                
                fig_status = px.pie(
                    status_counts_p,
                    values="contagem",
                    names="status",
                    hole=0.4,
                    color_discrete_sequence=px.colors.qualitative.Pastel
                )
                fig_status.update_traces(
                    textposition="inside",
                    textinfo="percent+label",
                    hovertemplate="<b>%{label}</b><br>%{value} contas (%{percent})"
                )
                fig_status.update_layout(
                    showlegend=False,
                    margin=dict(l=20, r=20, t=30, b=20),
                    height=350
                )
                
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.plotly_chart(fig_status, use_container_width=True)
                
                with col2:
                    st.markdown("""
                    <div style="background: #f8f9fa; padding: 1rem; border-radius: 10px;">
                        <h4 style="margin-top: 0;">Legenda</h4>
                        <div style="display: flex; align-items: center; margin-bottom: 8px;">
                            <div style="width: 12px; height: 12px; background: #636EFA; border-radius: 50%; margin-right: 8px;"></div>
                            <span>Em Aberto</span>
                        </div>
                        <div style="display: flex; align-items: center; margin-bottom: 8px;">
                            <div style="width: 12px; height: 12px; background: #EF553B; border-radius: 50%; margin-right: 8px;"></div>
                            <span>Em Atraso</span>
                        </div>
                        <div style="display: flex; align-items: center; margin-bottom: 8px;">
                            <div style="width: 12px; height: 12px; background: #00CC96; border-radius: 50%; margin-right: 8px;"></div>
                            <span>Pago</span>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                
                st.markdown("---")
                
                # Evolução mensal
                st.markdown("#### 📈 Evolução Mensal")
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
                
                fig_evolucao = go.Figure()
                fig_evolucao.add_trace(go.Scatter(
                    x=monthly_group_p["mes_ano_str"],
                    y=monthly_group_p["total_mes"],
                    name="Total",
                    line=dict(color="#6e8efb", width=3),
                    mode="lines+markers",
                    hovertemplate="<b>%{x}</b><br>Total: R$ %{y:,.2f}<extra></extra>"
                ))
                fig_evolucao.add_trace(go.Scatter(
                    x=monthly_group_p["mes_ano_str"],
                    y=monthly_group_p["pagos_mes"],
                    name="Pagas",
                    line=dict(color="#00CC96", width=2),
                    mode="lines+markers",
                    hovertemplate="<b>%{x}</b><br>Pagas: R$ %{y:,.2f}<extra></extra>"
                ))
                fig_evolucao.add_trace(go.Scatter(
                    x=monthly_group_p["mes_ano_str"],
                    y=monthly_group_p["pendentes_mes"],
                    name="Pendentes",
                    line=dict(color="#EF553B", width=2),
                    mode="lines+markers",
                    hovertemplate="<b>%{x}</b><br>Pendentes: R$ %{y:,.2f}<extra></extra>"
                ))
                
                fig_evolucao.update_layout(
                    hovermode="x unified",
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=1.02,
                        xanchor="right",
                        x=1
                    ),
                    margin=dict(l=20, r=20, t=30, b=20),
                    height=400,
                    xaxis_title="Mês/Ano",
                    yaxis_title="Valor (R$)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    paper_bgcolor="rgba(0,0,0,0)"
                )
                
                st.plotly_chart(fig_evolucao, use_container_width=True)
                
                # Top 10 fornecedores
                st.markdown("---")
                st.markdown("#### 🏆 Top 10 Fornecedores")
                top_fornecedores = (
                    df_all_p.groupby("fornecedor")
                    .agg(total=("valor", "sum"), contagem=("valor", "count"))
                    .sort_values("total", ascending=False)
                    .head(10)
                    .reset_index()
                )
                
                fig_fornecedores = px.bar(
                    top_fornecedores,
                    x="total",
                    y="fornecedor",
                    orientation="h",
                    color="contagem",
                    color_continuous_scale="Blues",
                    labels={"total": "Valor Total (R$)", "fornecedor": "", "contagem": "Nº Contas"},
                    hover_data={"contagem": True}
                )
                fig_fornecedores.update_layout(
                    height=500,
                    xaxis_title="Valor Total (R$)",
                    yaxis_title="",
                    yaxis={"categoryorder": "total ascending"},
                    margin=dict(l=20, r=20, t=30, b=20),
                    coloraxis_colorbar=dict(title="Nº Contas")
                )
                
                st.plotly_chart(fig_fornecedores, use_container_width=True)
                
                # Download dos dados
                st.markdown("---")
                with st.expander("💾 Exportar Dados", expanded=False):
                    try:
                        with open(EXCEL_PAGAR, "rb") as f:
                            st.download_button(
                                label="Baixar Planilha Completa (Contas a Pagar)",
                                data=f.read(),
                                file_name=EXCEL_PAGAR,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    except Exception as e:
                        st.error(f"Erro ao preparar download: {e}")

    with tab2:
        if not sheets_r:
            st.warning("Nenhuma aba válida encontrada em Contas a Receber")
        else:
            df_all_r = pd.concat([load_data(EXCEL_RECEBER, s) for s in sheets_r], ignore_index=True)
            
            if df_all_r.empty:
                st.info("Nenhum dado encontrado nas planilhas de Contas a Receber")
            else:
                # Métricas principais
                total_r = df_all_r["valor"].sum()
                num_lanc_r = len(df_all_r)
                media_r = df_all_r["valor"].mean() if num_lanc_r else 0
                atrasados_r = df_all_r[df_all_r["status_pagamento"] == "Em Atraso"]
                num_atras_r = len(atrasados_r)
                perc_atras_r = (num_atras_r / num_lanc_r * 100) if num_lanc_r else 0
                
                # Layout de métricas
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">Total a Receber</div>
                        <div class="metric-value">R$ {total_r:,.2f}</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">Lançamentos</div>
                        <div class="metric-value">{num_lanc_r}</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col3:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">Média por Conta</div>
                        <div class="metric-value">R$ {media_r:,.2f}</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col4:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">Em Atraso</div>
                        <div class="metric-value">{perc_atras_r:.1f}%</div>
                        <div style="font-size: 0.8rem; color: {'#e74c3c' if perc_atras_r > 10 else '#27ae60'}">
                            ({num_atras_r} conta{'s' if num_atras_r != 1 else ''})
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                
                st.markdown("---")
                
                # Gráfico de distribuição por status
                st.markdown("#### 📊 Distribuição por Status")
                status_counts_r = (
                    df_all_r["status_pagamento"]
                    .value_counts()
                    .rename_axis("status")
                    .reset_index(name="contagem")
                )
                
                fig_status = px.pie(
                    status_counts_r,
                    values="contagem",
                    names="status",
                    hole=0.4,
                    color_discrete_sequence=px.colors.qualitative.Pastel
                )
                fig_status.update_traces(
                    textposition="inside",
                    textinfo="percent+label",
                    hovertemplate="<b>%{label}</b><br>%{value} contas (%{percent})"
                )
                fig_status.update_layout(
                    showlegend=False,
                    margin=dict(l=20, r=20, t=30, b=20),
                    height=350
                )
                
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.plotly_chart(fig_status, use_container_width=True)
                
                with col2:
                    st.markdown("""
                    <div style="background: #f8f9fa; padding: 1rem; border-radius: 10px;">
                        <h4 style="margin-top: 0;">Legenda</h4>
                        <div style="display: flex; align-items: center; margin-bottom: 8px;">
                            <div style="width: 12px; height: 12px; background: #636EFA; border-radius: 50%; margin-right: 8px;"></div>
                            <span>A Receber</span>
                        </div>
                        <div style="display: flex; align-items: center; margin-bottom: 8px;">
                            <div style="width: 12px; height: 12px; background: #EF553B; border-radius: 50%; margin-right: 8px;"></div>
                            <span>Em Atraso</span>
                        </div>
                        <div style="display: flex; align-items: center; margin-bottom: 8px;">
                            <div style="width: 12px; height: 12px; background: #00CC96; border-radius: 50%; margin-right: 8px;"></div>
                            <span>Recebido</span>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                
                st.markdown("---")
                
                # Evolução mensal
                st.markdown("#### 📈 Evolução Mensal")
                df_all_r["mes_ano"] = df_all_r["vencimento"].dt.to_period("M")
                monthly_group_r = (
                    df_all_r
                    .groupby("mes_ano")
                    .agg(
                        total_mes=("valor", "sum"),
                        recebidos_mes=("valor", lambda x: x[df_all_r.loc[x.index, "status_pagamento"] == "Recebido"].sum()),
                        pendentes_mes=("valor", lambda x: x[df_all_r.loc[x.index, "status_pagamento"] != "Recebido"].sum())
                    )
                    .reset_index()
                )
                monthly_group_r["mes_ano_str"] = monthly_group_r["mes_ano"].dt.strftime("%b/%Y")
                
                fig_evolucao = go.Figure()
                fig_evolucao.add_trace(go.Scatter(
                    x=monthly_group_r["mes_ano_str"],
                    y=monthly_group_r["total_mes"],
                    name="Total",
                    line=dict(color="#6e8efb", width=3),
                    mode="lines+markers",
                    hovertemplate="<b>%{x}</b><br>Total: R$ %{y:,.2f}<extra></extra>"
                ))
                fig_evolucao.add_trace(go.Scatter(
                    x=monthly_group_r["mes_ano_str"],
                    y=monthly_group_r["recebidos_mes"],
                    name="Recebidos",
                    line=dict(color="#00CC96", width=2),
                    mode="lines+markers",
                    hovertemplate="<b>%{x}</b><br>Recebidos: R$ %{y:,.2f}<extra></extra>"
                ))
                fig_evolucao.add_trace(go.Scatter(
                    x=monthly_group_r["mes_ano_str"],
                    y=monthly_group_r["pendentes_mes"],
                    name="Pendentes",
                    line=dict(color="#EF553B", width=2),
                    mode="lines+markers",
                    hovertemplate="<b>%{x}</b><br>Pendentes: R$ %{y:,.2f}<extra></extra>"
                ))
                
                fig_evolucao.update_layout(
                    hovermode="x unified",
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=1.02,
                        xanchor="right",
                        x=1
                    ),
                    margin=dict(l=20, r=20, t=30, b=20),
                    height=400,
                    xaxis_title="Mês/Ano",
                    yaxis_title="Valor (R$)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    paper_bgcolor="rgba(0,0,0,0)"
                )
                
                st.plotly_chart(fig_evolucao, use_container_width=True)
                
                # Top 10 clientes
                st.markdown("---")
                st.markdown("#### 🏆 Top 10 Clientes")
                top_clientes = (
                    df_all_r.groupby("fornecedor")
                    .agg(total=("valor", "sum"), contagem=("valor", "count"))
                    .sort_values("total", ascending=False)
                    .head(10)
                    .reset_index()
                )
                
                fig_clientes = px.bar(
                    top_clientes,
                    x="total",
                    y="fornecedor",
                    orientation="h",
                    color="contagem",
                    color_continuous_scale="Blues",
                    labels={"total": "Valor Total (R$)", "fornecedor": "", "contagem": "Nº Contas"},
                    hover_data={"contagem": True}
                )
                fig_clientes.update_layout(
                    height=500,
                    xaxis_title="Valor Total (R$)",
                    yaxis_title="",
                    yaxis={"categoryorder": "total ascending"},
                    margin=dict(l=20, r=20, t=30, b=20),
                    coloraxis_colorbar=dict(title="Nº Contas")
                )
                
                st.plotly_chart(fig_clientes, use_container_width=True)
                
                # Download dos dados
                st.markdown("---")
                with st.expander("💾 Exportar Dados", expanded=False):
                    try:
                        with open(EXCEL_RECEBER, "rb") as f:
                            st.download_button(
                                label="Baixar Planilha Completa (Contas a Receber)",
                                data=f.read(),
                                file_name=EXCEL_RECEBER,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    except Exception as e:
                        st.error(f"Erro ao preparar download: {e}")


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
        df_display = df[df["status_pagamento"] == "Pago"].copy()
    elif view_sel == "Pendentes":
        df_display = df[df["status_pagamento"] != "Pago"].copy()
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
    
    if forn != "Todos":
        df_display = df_display[df_display["fornecedor"] == forn]
    if status_sel != "Todos":
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
        if not df_display.empty:
            idx = st.number_input(
                "Índice da linha (baseado na lista acima):",
                min_value=0,
                max_value=len(df_display) - 1,
                step=1,
                key="edit_pagar"
            )
            
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
                
                if save_data(EXCEL_PAGAR, aba, df):
                    st.success("Registro atualizado com sucesso!")
                    df = load_data(EXCEL_PAGAR, aba)
                    
                    # Reaplica filtros
                    if view_sel == "Pagas":
                        df_display = df[df["status_pagamento"] == "Pago"].copy()
                    elif view_sel == "Pendentes":
                        df_display = df[df["status_pagamento"] != "Pago"].copy()
                    else:
                        df_display = df.copy()
                    
                    if forn != "Todos":
                        df_display = df_display[df_display["fornecedor"] == forn]
                    if status_sel != "Todos":
                        df_display = df_display[df_display["estado"] == status_sel]
                    
                    table_placeholder.dataframe(df_display[cols_para_exibir], height=250)
                else:
                    st.error("Erro ao salvar alterações.")

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
                rec_rem = df_display.iloc[idx_rem]
                orig_idx = rec_rem.name
                
                try:
                    wb = load_workbook(EXCEL_PAGAR)
                    ws = wb[aba]
                    header_row = 8
                    excel_row = header_row + 1 + orig_idx
                    
                    headers = [
                        str(ws.cell(row=header_row, column=col).value).strip().lower()
                        for col in range(2, ws.max_column + 1)
                    ]
                    
                    field_map = {
                        "data_nf": ["data_nf", "data documento", "data da nota fiscal"],
                        "forma_pagamento": ["forma_pagamento", "descrição"],
                        "fornecedor": ["fornecedor"],
                        "os": ["os", "documento"],
                        "vencimento": ["vencimento"],
                        "valor": ["valor"],
                        "estado": ["estado"],
                        "boleto": ["boleto"],
                        "comprovante": ["comprovante"]
                    }
                    
                    cols_to_clear = []
                    for key, names in field_map.items():
                        for i, h in enumerate(headers):
                            if h in names:
                                cols_to_clear.append(i + 2)
                                break

                    for col in cols_to_clear:
                        ws.cell(row=excel_row, column=col, value=None)
                    
                    wb.save(EXCEL_PAGAR)
                    st.success("Registro removido com sucesso!")
                    
                    # Recarrega dados
                    df = load_data(EXCEL_PAGAR, aba)
                    
                    # Reaplica filtros
                    if view_sel == "Pagas":
                        df_display = df[df["status_pagamento"] == "Pago"].copy()
                    elif view_sel == "Pendentes":
                        df_display = df[df["status_pagamento"] != "Pago"].copy()
                    else:
                        df_display = df.copy()
                    
                    if forn != "Todos":
                        df_display = df_display[df_display["fornecedor"] == forn]
                    if status_sel != "Todos":
                        df_display = df_display[df_display["estado"] == status_sel]
                    
                    table_placeholder.dataframe(df_display[cols_para_exibir], height=250)
                    
                except Exception as e:
                    st.error(f"Erro ao remover registro: {e}")

    with st.expander("📎 Anexar Documentos"):
        if not df_display.empty:
            idx2 = st.number_input(
                "Índice para anexar (baseado na lista acima):",
                min_value=0, 
                max_value=len(df_display) - 1, 
                step=1, 
                key="idx_anex_pagar"
            )
            
            rec_anex = df_display.iloc[idx2]
            orig_idx_anex_candidates = df[
                (df["fornecedor"] == rec_anex["fornecedor"]) &
                (df["valor"] == rec_anex["valor"]) &
                (df["vencimento"] == rec_anex["vencimento"])
            ].index
            orig_idx_anex = orig_idx_anex_candidates[0] if len(orig_idx_anex_candidates) > 0 else rec_anex.name
            
            uploaded = st.file_uploader(
                "Selecione (pdf/jpg/png):", 
                type=["pdf", "jpg", "png"], 
                key=f"up_pagar_{aba}_{idx2}"
            )
            
            if uploaded:
                destino = os.path.join(
                    ANEXOS_DIR, 
                    "Contas a Pagar", 
                    f"Pagar_{aba}_{orig_idx_anex}_{uploaded.name}"
                )
                with open(destino, "wb") as f:
                    f.write(uploaded.getbuffer())
                st.success(f"Documento salvo em: {destino}")

    with st.expander("➕ Adicionar Nova Conta"):
        coln1, coln2 = st.columns(2)
        with coln1:
            data_nf = st.date_input(
                "Data N/F:",
                value=date.today(),
                key="nova_data_nf_pagar"
            )
            forma_pag = st.text_input(
                "Descrição:",
                key="nova_descricao_pagar"
            )
            forn_new = st.text_input(
                "Fornecedor:",
                key="novo_fornecedor_pagar"
            )
        with coln2:
            os_new = st.text_input(
                "Documento/OS:",
                key="novo_os_pagar"
            )
            venc_new = st.date_input(
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

        estado_opt = ["Em Aberto", "Pago"]
        situ_opt = ["Em Atraso", "Pago", "Em Aberto"]
        estado_new = st.selectbox(
            "Estado:",
            options=estado_opt,
            key="estado_novo_pagar"
        )
        situ_new = st.selectbox(
            "Situação:",
            options=situ_opt,
            key="situacao_novo_pagar"
        )

        boleto_file = st.file_uploader(
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

            if add_record(EXCEL_PAGAR, aba, record):
                st.success("Nova conta adicionada com sucesso!")
                
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

                # Recarrega dados
                df = load_data(EXCEL_PAGAR, aba)
                cols_show = ["data_nf", "fornecedor", "valor", "vencimento", "estado", "status_pagamento"]
                cols_to_display = [c for c in cols_show if c in df.columns]
                table_placeholder.dataframe(df[cols_to_display], height=250)
            else:
                st.error("Erro ao adicionar nova conta.")

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
    
    if not os.path.isfile(EXCEL_RECEBER):
        st.error(f"Arquivo '{EXCEL_RECEBER}' não encontrado. Verifique o caminho.")
        st.stop()
    
    existing = get_existing_sheets(EXCEL_RECEBER)
    aba = st.selectbox(
        "Selecione o mês:",
        FULL_MONTHS,
        index=FULL_MONTHS.index(date.today().strftime("%m"))
    )
    df = load_data(EXCEL_RECEBER, aba)
    
    if df.empty:
        st.info("Nenhum registro encontrado para este mês (ou a aba não existia).")
    
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
    
    with st.expander("🔍 Filtros"):
        col1, col2 = st.columns(2)
        with col1:
            forn = st.selectbox(
                "Cliente",
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
        cols_show = ["data_nf", "cliente", "valor", "vencimento", "estado", "status_pagamento"]
        cols_to_display = [c for c in cols_show if c in df_display.columns]
        table_placeholder_r = st.empty()
        table_placeholder_r.dataframe(df_display[cols_to_display], height=250)
    
    st.markdown("---")

    with st.expander("✏️ Editar Registro"):
        if df_display.empty:
            st.info("Nenhum registro para editar.")
        else:
            idx = st.number_input(
                "Índice da linha (baseado na lista acima):",
                min_value=0,
                max_value=len(df_display) - 1,
                step=1,
                key="edit_receber"
            )
            
            rec = df_display.iloc[idx]
            orig_idx = rec.name
            
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
                situ_opt = ["Em Atraso", "Recebido", "A Receber"]
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
            
            if st.button("💾 Salvar Alterações", key="salvar_receber"):
                df.at[orig_idx, "valor"] = new_val
                df.at[orig_idx, "vencimento"] = pd.to_datetime(new_venc)
                df.at[orig_idx, "estado"] = new_estado
                df.at[orig_idx, "situacao"] = new_sit
                
                if save_data(EXCEL_RECEBER, aba, df):
                    st.success("Registro atualizado com sucesso!")
                    df = load_data(EXCEL_RECEBER, aba)
                    
                    # Reaplica filtros
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
                    
                    table_placeholder_r.dataframe(df_display[cols_to_display], height=250)
                else:
                    st.error("Erro ao salvar alterações.")

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
                orig_idx = rec.name
                
                try:
                    wb = load_workbook(EXCEL_RECEBER)
                    ws = wb[aba]
                    ws.delete_rows(8 + 1 + orig_idx)
                    wb.save(EXCEL_RECEBER)
                    st.success("Registro removido com sucesso!")
                    
                    # Recarrega dados
                    df = load_data(EXCEL_RECEBER, aba)
                    
                    # Reaplica filtros
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
                    
                    table_placeholder_r.dataframe(df_display[cols_to_display], height=250)
                    
                except Exception as e:
                    st.error(f"Erro ao remover registro: {e}")

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
            orig2 = rec2.name
            
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

    with st.expander("➕ Adicionar Nova Conta"):
        coln1, coln2 = st.columns(2)
        with coln1:
            data_nf = st.date_input("Data N/F:", value=date.today(), key="nova_data_nf_receber")
            forma_pag = st.text_input("Descrição:", key="nova_descricao_receber")
            forn_new = st.text_input("Fornecedor:", key="novo_fornecedor_receber")
        with coln2:
            os_new = st.text_input("Documento/OS:", key="novo_os_receber")
            venc_new = st.date_input("Data de Vencimento:", value=date.today(), key="novo_venc_receber")
            valor_new = st.number_input("Valor (R$):", min_value=0.0, format="%.2f", key="novo_valor_receber2")

        estado_opt = ["A Receber", "Recebido"]
        situ_opt = ["Em Atraso", "Recebido", "A Receber"]
        estado_new = st.selectbox("Estado:", options=estado_opt, key="estado_novo_receber")
        situ_new = st.selectbox("Situação:", options=situ_opt, key="situacao_novo_receber")
        
        boleto_file = st.file_uploader("Boleto (opcional):", type=["pdf","jpg","png"], key="boleto_novo_receber")
        comprov_file = st.file_uploader("Comprovante (opcional):", type=["pdf","jpg","png"], key="comprov_novo_receber")

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
            
            if add_record(EXCEL_RECEBER, aba, record):
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
                
                # Recarrega dados
                df = load_data(EXCEL_RECEBER, aba)
                cols_to_display = [c for c in cols_show if c in df.columns]
                table_placeholder_r.dataframe(df[cols_to_display], height=250)
            else:
                st.error("Erro ao adicionar nova conta.")

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
