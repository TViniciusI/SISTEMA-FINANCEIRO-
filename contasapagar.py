import streamlit as st
import pandas as pd
import os
from datetime import datetime, date
from openpyxl import load_workbook
import plotly.express as px
from typing import Dict, List, Optional, Tuple

# Configurações iniciais
st.set_page_config(
    page_title="💼 Sistema Financeiro 2025",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Constantes
VALID_USERS = {
    "Vinicius": "vinicius4223",
    "Flavio": "1234",
}
EXCEL_PAGAR = "Contas a pagar 2025.xlsx"
EXCEL_RECEBER = "Contas a receber 2025.xlsx"
ANEXOS_DIR = "anexos"
FULL_MONTHS = [f"{i:02d}" for i in range(1, 13)]
COLUNAS_PADRAO = [
    "data_nf", "forma_pagamento", "fornecedor", "os",
    "vencimento", "valor", "estado", "situacao", "boleto", "comprovante"
]

# Funções auxiliares
def criar_pastas_necessarias():
    """Garante que todas as pastas necessárias existam."""
    for pasta in ["Contas a Pagar", "Contas a Receber"]:
        os.makedirs(os.path.join(ANEXOS_DIR, pasta), exist_ok=True)

def verificar_arquivos_excel():
    """Verifica se os arquivos Excel existem, criando se necessário."""
    for arquivo in [EXCEL_PAGAR, EXCEL_RECEBER]:
        if not os.path.isfile(arquivo):
            pd.DataFrame(columns=COLUNAS_PADRAO).to_excel(arquivo, index=False)

def check_login(username: str, password: str) -> bool:
    """Verifica as credenciais do usuário."""
    return VALID_USERS.get(username) == password

def get_existing_sheets(excel_path: str) -> List[str]:
    """Obtém todas as abas numéricas existentes no arquivo Excel."""
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
    except Exception:
        return []

def load_data(excel_path: str, sheet_name: str) -> pd.DataFrame:
    """Carrega os dados de uma aba específica do Excel."""
    if not os.path.isfile(excel_path):
        return pd.DataFrame(columns=COLUNAS_PADRAO + ["status_pagamento"])

    # Mapeamento de nomes de colunas
    col_mapping = {
        "data_nf": ["data documento", "data_nf", "data n/f", "data da nota fiscal"],
        "forma_pagamento": ["descrição", "forma_pagamento", "forma de pagamento"],
        "fornecedor": ["fornecedor", "cliente"],
        "os": ["documento", "os", "os interna"],
        "vencimento": ["vencimento"],
        "valor": ["valor"],
        "estado": ["estado"],
        "situacao": ["situação", "situacao"],
        "boleto": ["boleto", "boleto anexo"],
        "comprovante": ["comprovante", "comprovante de pagto"]
    }

    try:
        # Encontra o nome real da aba (pode ser "4" em vez de "04")
        sheet_lookup = {}
        with pd.ExcelFile(excel_path) as wb:
            for s in wb.sheet_names:
                nome = s.strip()
                if nome.lower() != "tutorial" and nome.isdigit():
                    sheet_lookup[f"{int(nome):02d}"] = nome
        
        if sheet_name not in sheet_lookup:
            return pd.DataFrame(columns=COLUNAS_PADRAO + ["status_pagamento"])
        
        real_sheet = sheet_lookup[sheet_name]
        df = pd.read_excel(excel_path, sheet_name=real_sheet, skiprows=7, header=0)
        
        # Renomeia colunas conforme mapeamento
        rename_map = {}
        for col in df.columns:
            col_normalizado = str(col).strip().lower()
            for target, aliases in col_mapping.items():
                for alias in aliases:
                    if alias.lower().strip() == col_normalizado:
                        rename_map[col] = target
                        break
        
        df = df.rename(columns=rename_map)
        
        # Garante colunas mínimas
        for col in ["fornecedor", "valor"]:
            if col not in df.columns:
                df[col] = pd.NA
        
        # Limpeza de dados
        df = df.dropna(subset=["fornecedor", "valor"], how="all").reset_index(drop=True)
        df["vencimento"] = pd.to_datetime(df["vencimento"], errors="coerce")
        df["valor"] = pd.to_numeric(df["valor"], errors="coerce")
        
        # Determina o status de pagamento
        is_receber = (excel_path == EXCEL_RECEBER)
        hoje = datetime.now().date()
        
        def determinar_status(row):
            estado = str(row.get("estado", "")).strip().lower()
            if estado == ("recebido" if is_receber else "pago"):
                return "Recebido" if is_receber else "Pago"
            
            data_venc = row["vencimento"].date() if pd.notna(row["vencimento"]) else None
            if data_venc:
                if data_venc < hoje:
                    return "Em Atraso"
                else:
                    return "A Receber" if is_receber else "Em Aberto"
            return "Sem Data"
        
        df["status_pagamento"] = df.apply(determinar_status, axis=1)
        return df
        
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return pd.DataFrame(columns=COLUNAS_PADRAO + ["status_pagamento"])

def save_data(excel_path: str, sheet_name: str, df: pd.DataFrame):
    """Salva os dados de volta no Excel, preservando formatação."""
    try:
        wb = load_workbook(excel_path)
        
        # Encontra o nome real da aba
        sheet_lookup = {}
        for s in wb.sheetnames:
            nome = s.strip()
            if nome.lower() != "tutorial" and nome.isdigit():
                sheet_lookup[f"{int(nome):02d}"] = nome
        
        real_sheet = sheet_lookup.get(sheet_name, sheet_name)
        
        if real_sheet not in wb.sheetnames:
            # Cria nova aba baseada na primeira aba numérica
            numeric_sheets = [s for s in wb.sheetnames if s.isdigit()]
            if numeric_sheets:
                template = wb[numeric_sheets[0]]
                new_sheet = wb.copy_worksheet(template)
                new_sheet.title = real_sheet
            else:
                new_sheet = wb.create_sheet(real_sheet)
            ws = new_sheet
        else:
            ws = wb[real_sheet]
        
        # Encontra cabeçalhos e mapeia colunas
        header_row = 8
        col_pos = {}
        for col in range(1, ws.max_column + 1):
            cell_val = str(ws.cell(row=header_row, column=col).value).strip().lower()
            if "data" in cell_val and ("nf" in cell_val or "nota" in cell_val):
                col_pos["data_nf"] = col
            elif "forma" in cell_val and "pagamento" in cell_val:
                col_pos["forma_pagamento"] = col
            elif "fornecedor" in cell_val or "cliente" in cell_val:
                col_pos["fornecedor"] = col
            elif "os" in cell_val or "documento" in cell_val:
                col_pos["os"] = col
            elif "vencimento" in cell_val:
                col_pos["vencimento"] = col
            elif "valor" in cell_val:
                col_pos["valor"] = col
            elif "estado" in cell_val:
                col_pos["estado"] = col
            elif "boleto" in cell_val:
                col_pos["boleto"] = col
            elif "comprovante" in cell_val:
                col_pos["comprovante"] = col
        
        # Atualiza os dados
        for idx, row in df.iterrows():
            excel_row = header_row + 1 + idx
            for col_name, col_num in col_pos.items():
                value = row.get(col_name)
                
                if pd.isna(value):
                    continue
                    
                if col_name in ["data_nf", "vencimento"]:
                    try:
                        value = pd.to_datetime(value).to_pydatetime()
                    except:
                        continue
                elif col_name == "valor":
                    try:
                        value = float(value)
                    except:
                        continue
                
                ws.cell(row=excel_row, column=col_num, value=value)
        
        wb.save(excel_path)
        return True
    except Exception as e:
        st.error(f"Erro ao salvar dados: {e}")
        return False

def add_record(excel_path: str, sheet_name: str, record: Dict):
    """Adiciona um novo registro ao arquivo Excel."""
    try:
        wb = load_workbook(excel_path)
        
        # Encontra o nome real da aba
        sheet_lookup = {}
        for s in wb.sheetnames:
            nome = s.strip()
            if nome.lower() != "tutorial" and nome.isdigit():
                sheet_lookup[f"{int(nome):02d}"] = nome
        
        real_sheet = sheet_lookup.get(sheet_name, sheet_name)
        
        if real_sheet not in wb.sheetnames:
            # Cria nova aba baseada na primeira aba numérica
            numeric_sheets = [s for s in wb.sheetnames if s.isdigit()]
            if numeric_sheets:
                template = wb[numeric_sheets[0]]
                new_sheet = wb.copy_worksheet(template)
                new_sheet.title = real_sheet
            else:
                new_sheet = wb.create_sheet(real_sheet)
            ws = new_sheet
        else:
            ws = wb[real_sheet]
        
        # Encontra cabeçalhos e mapeia colunas
        header_row = 8
        col_pos = {}
        for col in range(1, ws.max_column + 1):
            cell_val = str(ws.cell(row=header_row, column=col).value).strip().lower()
            if "data" in cell_val and ("nf" in cell_val or "nota" in cell_val):
                col_pos["data_nf"] = col
            elif "forma" in cell_val and "pagamento" in cell_val:
                col_pos["forma_pagamento"] = col
            elif "fornecedor" in cell_val or "cliente" in cell_val:
                col_pos["fornecedor"] = col
            elif "os" in cell_val or "documento" in cell_val:
                col_pos["os"] = col
            elif "vencimento" in cell_val:
                col_pos["vencimento"] = col
            elif "valor" in cell_val:
                col_pos["valor"] = col
            elif "estado" in cell_val:
                col_pos["estado"] = col
            elif "boleto" in cell_val:
                col_pos["boleto"] = col
            elif "comprovante" in cell_val:
                col_pos["comprovante"] = col
        
        # Encontra primeira linha vazia
        next_row = header_row + 1
        while next_row <= ws.max_row and ws.cell(row=next_row, column=col_pos.get("fornecedor", 2)).value:
            next_row += 1
        
        # Preenche os valores
        for col_name, col_num in col_pos.items():
            value = record.get(col_name)
            if value is None:
                continue
                
            if col_name in ["data_nf", "vencimento"]:
                try:
                    value = pd.to_datetime(value).to_pydatetime()
                except:
                    continue
            elif col_name == "valor":
                try:
                    value = float(value)
                except:
                    continue
            
            ws.cell(row=next_row, column=col_num, value=value)
        
        wb.save(excel_path)
        return True
    except Exception as e:
        st.error(f"Erro ao adicionar registro: {e}")
        return False

def remove_record(excel_path: str, sheet_name: str, index: int) -> bool:
    """Remove um registro específico do arquivo Excel."""
    try:
        wb = load_workbook(excel_path)

        # Encontra o nome real da aba
        sheet_lookup = {}
        for s in wb.sheetnames:
            nome = s.strip()
            if nome.lower() != "tutorial" and nome.isdigit():
                sheet_lookup[f"{int(nome):02d}"] = nome

        real_sheet = sheet_lookup.get(sheet_name, sheet_name)

        if real_sheet not in wb.sheetnames:
            return False

        ws = wb[real_sheet]
        header_row = 8
        row_to_delete = header_row + 1 + index

        if row_to_delete > ws.max_row:
            return False

        ws.delete_rows(row_to_delete)
        wb.save(excel_path)
        return True

    except Exception as e:
        st.error(f"Erro ao remover registro: {e}")
        return False

def format_currency(value: float) -> str:
    """Formata um valor como moeda brasileira."""
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def display_dashboard():
    """Exibe o painel de controle com métricas e gráficos."""
    st.subheader("📊 Painel de Controle Financeiro Avançado")
    
    # Verifica arquivos
    if not os.path.isfile(EXCEL_PAGAR):
        st.error(f"Arquivo '{EXCEL_PAGAR}' não encontrado.")
        return
    if not os.path.isfile(EXCEL_RECEBER):
        st.error(f"Arquivo '{EXCEL_RECEBER}' não encontrado.")
        return
    
    # Carrega dados
    sheets_p = get_existing_sheets(EXCEL_PAGAR)
    sheets_r = get_existing_sheets(EXCEL_RECEBER)
    
    if not sheets_p and not sheets_r:
        st.warning("Nenhuma aba numérica encontrada nos arquivos.")
        return
    
    # Abas para Pagar e Receber
    tab1, tab2 = st.tabs(["📥 Contas a Pagar", "📤 Contas a Receber"])
    
    with tab1:
        if not sheets_p:
            st.warning("Nenhuma aba válida encontrada em 'Contas a Pagar'.")
        else:
            # Carrega todos os dados de pagar
            df_all_p = pd.concat([load_data(EXCEL_PAGAR, s) for s in sheets_p], ignore_index=True)
            
            # Métricas
            total_p = df_all_p["valor"].sum()
            num_lanc_p = len(df_all_p)
            media_p = df_all_p["valor"].mean() if num_lanc_p else 0
            atrasados_p = df_all_p[df_all_p["status_pagamento"] == "Em Atraso"]
            num_atras_p = len(atrasados_p)
            perc_atras_p = (num_atras_p / num_lanc_p * 100) if num_lanc_p else 0
            
            # Exibe métricas
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total a Pagar", format_currency(total_p))
            col2.metric("Nº Lançamentos", num_lanc_p)
            col3.metric("Média por Conta", format_currency(media_p))
            col4.metric("Em Atraso", f"{num_atras_p} ({perc_atras_p:.1f}%)")
            
            # Gráficos
            st.markdown("---")
            st.subheader("📈 Evolução Mensal")
            
            # Prepara dados mensais
            df_all_p["mes_ano"] = df_all_p["vencimento"].dt.to_period("M")
            monthly_p = df_all_p.groupby("mes_ano").agg(
                Total=("valor", "sum"),
                Pagas=("valor", lambda x: x[df_all_p.loc[x.index, "status_pagamento"] == "Pago"].sum()),
                Pendentes=("valor", lambda x: x[df_all_p.loc[x.index, "status_pagamento"] != "Pago"].sum())
            ).reset_index()
            
            monthly_p["Mês"] = monthly_p["mes_ano"].dt.strftime("%b/%Y")
            
            # Gráfico de linhas
            fig = px.line(
                monthly_p, 
                x="Mês", 
                y=["Total", "Pagas", "Pendentes"],
                title="Evolução de Contas a Pagar",
                labels={"value": "Valor (R$)", "variable": "Status"}
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # Gráfico de pizza por status
            status_counts_p = df_all_p["status_pagamento"].value_counts().reset_index()
            status_counts_p.columns = ["Status", "Quantidade"]
            
            fig = px.pie(
                status_counts_p,
                values="Quantidade",
                names="Status",
                title="Distribuição por Status",
                hole=0.3
            )
            st.plotly_chart(fig, use_container_width=True)
    
    with tab2:
        if not sheets_r:
            st.warning("Nenhuma aba válida encontrada em 'Contas a Receber'.")
        else:
            # Carrega todos os dados de receber
            df_all_r = pd.concat([load_data(EXCEL_RECEBER, s) for s in sheets_r], ignore_index=True)
            
            # Métricas
            total_r = df_all_r["valor"].sum()
            num_lanc_r = len(df_all_r)
            media_r = df_all_r["valor"].mean() if num_lanc_r else 0
            atrasados_r = df_all_r[df_all_r["status_pagamento"] == "Em Atraso"]
            num_atras_r = len(atrasados_r)
            perc_atras_r = (num_atras_r / num_lanc_r * 100) if num_lanc_r else 0
            
            # Exibe métricas
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total a Receber", format_currency(total_r))
            col2.metric("Nº Lançamentos", num_lanc_r)
            col3.metric("Média por Conta", format_currency(media_r))
            col4.metric("Em Atraso", f"{num_atras_r} ({perc_atras_r:.1f}%)")
            
            # Gráficos
            st.markdown("---")
            st.subheader("📈 Evolução Mensal")
            
            # Prepara dados mensais
            df_all_r["mes_ano"] = df_all_r["vencimento"].dt.to_period("M")
            monthly_r = df_all_r.groupby("mes_ano").agg(
                Total=("valor", "sum"),
                Recebidas=("valor", lambda x: x[df_all_r.loc[x.index, "status_pagamento"] == "Recebido"].sum()),
                Pendentes=("valor", lambda x: x[df_all_r.loc[x.index, "status_pagamento"] != "Recebido"].sum())
            ).reset_index()
            
            monthly_r["Mês"] = monthly_r["mes_ano"].dt.strftime("%b/%Y")
            
            # Gráfico de linhas
            fig = px.line(
                monthly_r, 
                x="Mês", 
                y=["Total", "Recebidas", "Pendentes"],
                title="Evolução de Contas a Receber",
                labels={"value": "Valor (R$)", "variable": "Status"}
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # Gráfico de pizza por status
            status_counts_r = df_all_r["status_pagamento"].value_counts().reset_index()
            status_counts_r.columns = ["Status", "Quantidade"]
            
            fig = px.pie(
                status_counts_r,
                values="Quantidade",
                names="Status",
                title="Distribuição por Status",
                hole=0.3
            )
            st.plotly_chart(fig, use_container_width=True)

def display_pagar():
    """Exibe a interface para gerenciar contas a pagar."""
    st.subheader("🗂️ Contas a Pagar")
    
    if not os.path.isfile(EXCEL_PAGAR):
        st.error(f"Arquivo '{EXCEL_PAGAR}' não encontrado.")
        return
    
    existing_sheets = get_existing_sheets(EXCEL_PAGAR)
    if not existing_sheets:
        st.warning("Nenhuma aba válida encontrada no arquivo.")
        return
    
    # Seletor de mês
    aba = st.selectbox("Selecione o mês:", FULL_MONTHS, index=int(datetime.now().strftime("%m"))-1)
    
    # Carrega dados
    df = load_data(EXCEL_PAGAR, aba)
    
    # Filtros
    view_sel = st.radio("Visualizar:", ["Todos", "Pagas", "Pendentes"], horizontal=True)
    
    if view_sel == "Pagas":
        df_display = df[df["status_pagamento"] == "Pago"].copy()
    elif view_sel == "Pendentes":
        df_display = df[df["status_pagamento"] != "Pago"].copy()
    else:
        df_display = df.copy()
    
    with st.expander("🔍 Filtros Avançados"):
        col1, col2 = st.columns(2)
        with col1:
            fornecedor = st.selectbox(
                "Fornecedor",
                ["Todos"] + sorted(df["fornecedor"].dropna().unique().tolist())
            )
        with col2:
            status = st.selectbox(
                "Status",
                ["Todos"] + sorted(df["status_pagamento"].dropna().unique().tolist())
            )
    
    if fornecedor != "Todos":
        df_display = df_display[df_display["fornecedor"] == fornecedor]
    if status != "Todos":
        df_display = df_display[df_display["status_pagamento"] == status]
    
    # Exibe tabela
    st.markdown("---")
    st.subheader("📋 Lançamentos")
    
    if df_display.empty:
        st.warning("Nenhum registro encontrado para os filtros selecionados.")
    else:
        cols_to_show = [c for c in COLUNAS_PADRAO if c in df_display.columns] + ["status_pagamento"]
        st.dataframe(
            df_display[cols_to_show],
            height=400,
            use_container_width=True
        )
    
    # Operações CRUD
    st.markdown("---")
    st.subheader("✏️ Operações")
    
    tab_edit, tab_add, tab_del = st.tabs(["Editar", "Adicionar", "Remover"])
    
    with tab_edit:
        if df_display.empty:
            st.info("Nenhum registro para editar.")
        else:
            idx = st.number_input(
                "Índice do registro:",
                min_value=0,
                max_value=len(df_display)-1,
                step=1,
                key="edit_idx_pagar"
            )
            
            rec = df_display.iloc[idx]
            orig_idx = df[df["fornecedor"] == rec["fornecedor"]].index[0]
            
            col1, col2 = st.columns(2)
            with col1:
                novo_valor = st.number_input(
                    "Valor:",
                    value=float(rec["valor"]),
                    key="novo_valor_pagar"
                )
                novo_venc = st.date_input(
                    "Vencimento:",
                    value=rec["vencimento"].date() if pd.notna(rec["vencimento"]) else date.today(),
                    key="novo_venc_pagar"
                )
            with col2:
                novo_estado = st.selectbox(
                    "Estado:",
                    options=["Em Aberto", "Pago"],
                    index=0 if rec["estado"] == "Em Aberto" else 1,
                    key="novo_estado_pagar"
                )
                novo_situacao = st.selectbox(
                    "Situação:",
                    options=["Em Atraso", "Pago", "Em Aberto"],
                    index=["Em Atraso", "Pago", "Em Aberto"].index(rec["situacao"]) if rec["situacao"] in ["Em Atraso", "Pago", "Em Aberto"] else 0,
                    key="novo_situacao_pagar"
                )
            
            if st.button("💾 Salvar Alterações", key="btn_save_pagar"):
                df.at[orig_idx, "valor"] = novo_valor
                df.at[orig_idx, "vencimento"] = novo_venc
                df.at[orig_idx, "estado"] = novo_estado
                df.at[orig_idx, "situacao"] = novo_situacao
                
                if save_data(EXCEL_PAGAR, aba, df):
                    st.success("Registro atualizado com sucesso!")
                    st.experimental_rerun()
    
    with tab_add:
        col1, col2 = st.columns(2)
        with col1:
            nova_data = st.date_input(
                "Data N/F:",
                value=date.today(),
                key="nova_data_pagar"
            )
            nova_desc = st.text_input(
                "Descrição:",
                key="nova_desc_pagar"
            )
            novo_forn = st.text_input(
                "Fornecedor:",
                key="novo_forn_pagar"
            )
        with col2:
            novo_os = st.text_input(
                "Documento/OS:",
                key="novo_os_pagar"
            )
            novo_venc = st.date_input(
                "Vencimento:",
                value=date.today(),
                key="novo_venc_add_pagar"
            )
            novo_valor = st.number_input(
                "Valor:",
                min_value=0.01,
                value=100.00,
                step=1.0,
                format="%.2f",
                key="novo_valor_add_pagar"
            )
        
        novo_estado = st.selectbox(
            "Estado:",
            options=["Em Aberto", "Pago"],
            index=0,
            key="novo_estado_add_pagar"
        )
        novo_situacao = st.selectbox(
            "Situação:",
            options=["Em Atraso", "Pago", "Em Aberto"],
            index=2,
            key="novo_situacao_add_pagar"
        )
        
        novo_boleto = st.file_uploader(
            "Boleto (opcional):",
            type=["pdf", "jpg", "png"],
            key="novo_boleto_pagar"
        )
        novo_comprov = st.file_uploader(
            "Comprovante (opcional):",
            type=["pdf", "jpg", "png"],
            key="novo_comprov_pagar"
        )
        
        if st.button("➕ Adicionar Registro", key="btn_add_pagar"):
            record = {
                "data_nf": nova_data,
                "forma_pagamento": nova_desc,
                "fornecedor": novo_forn,
                "os": novo_os,
                "vencimento": novo_venc,
                "valor": novo_valor,
                "estado": novo_estado,
                "situacao": novo_situacao,
                "boleto": "",
                "comprovante": ""
            }
            
            # Salva anexos
            if novo_boleto:
                boleto_path = os.path.join(
                    ANEXOS_DIR,
                    "Contas a Pagar",
                    f"Pagar_{aba}_{novo_forn}_{novo_boleto.name}"
                )
                with open(boleto_path, "wb") as f:
                    f.write(novo_boleto.getbuffer())
                record["boleto"] = boleto_path
            
            if novo_comprov:
                comprov_path = os.path.join(
                    ANEXOS_DIR,
                    "Contas a Pagar",
                    f"Pagar_{aba}_{novo_forn}_{novo_comprov.name}"
                )
                with open(comprov_path, "wb") as f:
                    f.write(novo_comprov.getbuffer())
                record["comprovante"] = comprov_path
            
            if add_record(EXCEL_PAGAR, aba, record):
                st.success("Registro adicionado com sucesso!")
                st.experimental_rerun()
    
    with tab_del:
        if df_display.empty:
            st.info("Nenhum registro para remover.")
        else:
            idx_del = st.number_input(
                "Índice do registro:",
                min_value=0,
                max_value=len(df_display)-1,
                step=1,
                key="del_idx_pagar"
            )
            
            rec = df_display.iloc[idx_del]
            st.warning(f"Você está prestes a remover o registro: {rec['fornecedor']} - {format_currency(rec['valor'])}")
            
            if st.button("🗑️ Confirmar Remoção", key="btn_del_pagar"):
                orig_idx = df[df["fornecedor"] == rec["fornecedor"]].index[0]
                if remove_record(EXCEL_PAGAR, aba, orig_idx):
                    st.success("Registro removido com sucesso!")
                    st.experimental_rerun()

def display_receber():
    """Exibe a interface para gerenciar contas a receber."""
    st.subheader("🗂️ Contas a Receber")
    
    if not os.path.isfile(EXCEL_RECEBER):
        st.error(f"Arquivo '{EXCEL_RECEBER}' não encontrado.")
        return
    
    existing_sheets = get_existing_sheets(EXCEL_RECEBER)
    if not existing_sheets:
        st.warning("Nenhuma aba válida encontrada no arquivo.")
        return
    
    # Seletor de mês
    aba = st.selectbox("Selecione o mês:", FULL_MONTHS, index=int(datetime.now().strftime("%m"))-1)
    
    # Carrega dados
    df = load_data(EXCEL_RECEBER, aba)
    
    # Filtros
    view_sel = st.radio("Visualizar:", ["Todos", "Recebidas", "Pendentes"], horizontal=True)
    
    if view_sel == "Recebidas":
        df_display = df[df["status_pagamento"] == "Recebido"].copy()
    elif view_sel == "Pendentes":
        df_display = df[df["status_pagamento"] != "Recebido"].copy()
    else:
        df_display = df.copy()
    
with st.expander("🔍 Filtros Avançados"):
    col1, col2 = st.columns(2)
    with col1:
        cliente = st.selectbox(
            "Cliente:",
            ["Todos"] + sorted(df["fornecedor"].dropna().unique().tolist())
        )
    with col2:
        status = st.selectbox(
            "Status:",
            ["Todos"] + sorted(df["status_pagamento"].dropna().unique().tolist())
        )

    if cliente != "Todos":
        df_display = df_display[df_display["fornecedor"] == cliente]
    if status != "Todos":
        df_display = df_display[df_display["status_pagamento"] == status]
    # Exibe tabela
    st.markdown("---")
    st.subheader("📋 Lançamentos")
    
    if df_display.empty:
        st.warning("Nenhum registro encontrado para os filtros selecionados.")
    else:
        cols_to_show = [c for c in COLUNAS_PADRAO if c in df_display.columns] + ["status_pagamento"]
        st.dataframe(
            df_display[cols_to_show],
            height=400,
            use_container_width=True
        )
    
    # Operações CRUD
    st.markdown("---")
    st.subheader("✏️ Operações")
    
    tab_edit, tab_add, tab_del = st.tabs(["Editar", "Adicionar", "Remover"])
    
    with tab_edit:
        if df_display.empty:
            st.info("Nenhum registro para editar.")
        else:
            idx = st.number_input(
                "Índice do registro:",
                min_value=0,
                max_value=len(df_display)-1,
                step=1,
                key="edit_idx_receber"
            )
            
            rec = df_display.iloc[idx]
            orig_idx = df[df["fornecedor"] == rec["fornecedor"]].index[0]
            
            col1, col2 = st.columns(2)
            with col1:
                novo_valor = st.number_input(
                    "Valor:",
                    value=float(rec["valor"]),
                    key="novo_valor_receber"
                )
                novo_venc = st.date_input(
                    "Vencimento:",
                    value=rec["vencimento"].date() if pd.notna(rec["vencimento"]) else date.today(),
                    key="novo_venc_receber"
                )
            with col2:
                novo_estado = st.selectbox(
                    "Estado:",
                    options=["A Receber", "Recebido"],
                    index=0 if rec["estado"] == "A Receber" else 1,
                    key="novo_estado_receber"
                )
                novo_situacao = st.selectbox(
                    "Situação:",
                    options=["Em Atraso", "Recebido", "A Receber"],
                    index=["Em Atraso", "Recebido", "A Receber"].index(rec["situacao"]) if rec["situacao"] in ["Em Atraso", "Recebido", "A Receber"] else 0,
                    key="novo_situacao_receber"
                )
            
            if st.button("💾 Salvar Alterações", key="btn_save_receber"):
                df.at[orig_idx, "valor"] = novo_valor
                df.at[orig_idx, "vencimento"] = novo_venc
                df.at[orig_idx, "estado"] = novo_estado
                df.at[orig_idx, "situacao"] = novo_situacao
                
                if save_data(EXCEL_RECEBER, aba, df):
                    st.success("Registro atualizado com sucesso!")
                    st.experimental_rerun()
    
    with tab_add:
        col1, col2 = st.columns(2)
        with col1:
                       nova_data = st.date_input(
                "Data N/F:",
                value=date.today(),
                key="nova_data_receber"
            )
            nova_desc = st.text_input(
                "Descrição:",
                key="nova_desc_receber"
            )
            novo_cliente = st.text_input(
                "Cliente:",
                key="novo_cliente_receber"
            )
        with col2:
            novo_os = st.text_input(
                "Documento/OS:",
                key="novo_os_receber"
            )
            novo_venc = st.date_input(
                "Vencimento:",
                value=date.today(),
                key="novo_venc_add_receber"
            )
            novo_valor = st.number_input(
                "Valor:",
                min_value=0.01,
                value=100.00,
                step=1.0,
                format="%.2f",
                key="novo_valor_add_receber"
            )
        
        novo_estado = st.selectbox(
            "Estado:",
            options=["A Receber", "Recebido"],
            index=0,
            key="novo_estado_add_receber"
        )
        novo_situacao = st.selectbox(
            "Situação:",
            options=["Em Atraso", "Recebido", "A Receber"],
            index=2,
            key="novo_situacao_add_receber"
        )
        
        novo_boleto = st.file_uploader(
            "Boleto (opcional):",
            type=["pdf", "jpg", "png"],
            key="novo_boleto_receber"
        )
        novo_comprov = st.file_uploader(
            "Comprovante (opcional):",
            type=["pdf", "jpg", "png"],
            key="novo_comprov_receber"
        )
        
        if st.button("➕ Adicionar Registro", key="btn_add_receber"):
            record = {
                "data_nf": nova_data,
                "forma_pagamento": nova_desc,
                "fornecedor": novo_cliente,
                "os": novo_os,
                "vencimento": novo_venc,
                "valor": novo_valor,
                "estado": novo_estado,
                "situacao": novo_situacao,
                "boleto": "",
                "comprovante": ""
            }
            
            # Salva anexos
            if novo_boleto:
                boleto_path = os.path.join(
                    ANEXOS_DIR,
                    "Contas a Receber",
                    f"Receber_{aba}_{novo_cliente}_{novo_boleto.name}"
                )
                with open(boleto_path, "wb") as f:
                    f.write(novo_boleto.getbuffer())
                record["boleto"] = boleto_path
            
            if novo_comprov:
                comprov_path = os.path.join(
                    ANEXOS_DIR,
                    "Contas a Receber",
                    f"Receber_{aba}_{novo_cliente}_{novo_comprov.name}"
                )
                with open(comprov_path, "wb") as f:
                    f.write(novo_comprov.getbuffer())
                record["comprovante"] = comprov_path
            
            if add_record(EXCEL_RECEBER, aba, record):
                st.success("Registro adicionado com sucesso!")
                st.experimental_rerun()
    
    with tab_del:
        if df_display.empty:
            st.info("Nenhum registro para remover.")
        else:
            idx_del = st.number_input(
                "Índice do registro:",
                min_value=0,
                max_value=len(df_display)-1,
                step=1,
                key="del_idx_receber"
            )
            
            rec = df_display.iloc[idx_del]
            st.warning(f"Você está prestes a remover o registro: {rec['fornecedor']} - {format_currency(rec['valor'])}")
            
            if st.button("🗑️ Confirmar Remoção", key="btn_del_receber"):
                orig_idx = df[df["fornecedor"] == rec["fornecedor"]].index[0]
                if remove_record(EXCEL_RECEBER, aba, orig_idx):
                    st.success("Registro removido com sucesso!")
                    st.experimental_rerun()

# Sistema de login
def login_section():
    """Exibe a seção de login e gerencia o estado da sessão."""
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False
        st.session_state.username = ""
    
    if not st.session_state.logged_in:
        st.write("\n" * 5)
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.title("🔒 Login")
            username = st.text_input("Usuário:")
            password = st.text_input("Senha:", type="password")
            
            if st.button("Entrar"):
                if check_login(username, password):
                    st.session_state.logged_in = True
                    st.session_state.username = username
                    st.experimental_rerun()
                else:
                    st.error("Usuário ou senha inválidos.")
        st.stop()
    
    # Logout button in sidebar
    st.sidebar.write(f"Logado como: **{st.session_state.username}**")
    if st.sidebar.button("🚪 Sair"):
        st.session_state.logged_in = False
        st.session_state.username = ""
        st.experimental_rerun()

# Menu principal
def main_menu():
    """Exibe o menu principal e gerencia a navegação."""
    st.sidebar.markdown("""
    ## 📂 Navegação  
    Selecione a seção desejada para visualizar e gerenciar  
    suas contas a pagar e receber.  
    """, unsafe_allow_html=True)
    
    page = st.sidebar.radio("", ["Dashboard", "Contas a Pagar", "Contas a Receber"], index=0)
    
    st.markdown("""
    <div style="text-align: center; color: #4B8BBE; margin-bottom: 10px;">
        <h1>💼 Sistema Financeiro 2025</h1>
        <p style="color: #555; font-size: 16px;">Dashboard avançado com estatísticas e gráficos interativos.</p>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")
    
    if page == "Dashboard":
        display_dashboard()
    elif page == "Contas a Pagar":
        display_pagar()
    elif page == "Contas a Receber":
        display_receber()
    
    st.markdown("""
    <div style="text-align: center; font-size:12px; color:gray; margin-top: 20px;">
        <p>© 2025 Desenvolvido por Vinicius Magalhães</p>
    </div>
    """, unsafe_allow_html=True)

# Inicialização do app
def main():
    """Função principal que orquestra a execução do aplicativo."""
    criar_pastas_necessarias()
    verificar_arquivos_excel()
    login_section()
    main_menu()

if __name__ == "__main__":
    main()
