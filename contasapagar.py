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

# AUTENTICAÇÃO
VALID_USERS = {"Vinicius":"vinicius4223","Flavio":"1234"}
def check_login(u,p): return VALID_USERS.get(u)==p

if "logged_in" not in st.session_state:
    st.session_state.logged_in=False
    st.session_state.username=""
if not st.session_state.logged_in:
    st.write("\n"*5)
    c1,c2,c3 = st.columns([1,2,1])
    with c2:
        st.title("🔒 Login")
        ui = st.text_input("Usuário:")
        pw = st.text_input("Senha:",type="password")
        if st.button("Entrar"):
            if check_login(ui,pw):
                st.session_state.logged_in=True
                st.session_state.username=ui
            else:
                st.error("Usuário ou senha inválidos.")
    st.stop()

logged_user = st.session_state.username
st.sidebar.write(f"Logado como: **{logged_user}**")

# CONSTANTES
EXCEL_PAGAR   = "Contas a pagar 2025 Sistema.xlsx"
EXCEL_RECEBER = "Contas a Receber 2025 Sistema.xlsx"
ANEXOS_DIR    = "anexos"
FULL_MONTHS   = [f"{i:02d}" for i in range(1,13)]
for pasta in ["Contas a Pagar","Contas a Receber"]:
    os.makedirs(os.path.join(ANEXOS_DIR,pasta),exist_ok=True)

# AUXILIARES
def get_existing_sheets(path):
    try:
        sheets = pd.ExcelFile(path).sheet_names
        return sorted([s for s in sheets if s.strip().isdigit()])
    except:
        return []

def load_data(path,month):
    cols = ["data_nf","forma_pagamento","fornecedor","os",
            "vencimento","valor","estado","situacao","boleto","comprovante"]
    if not os.path.isfile(path) or month not in get_existing_sheets(path):
        return pd.DataFrame(columns=cols+["status_pagamento"])
    df = pd.read_excel(path,month,skiprows=7,header=0)
    df.rename(columns=lambda c:str(c).strip().lower(),inplace=True)
    df.rename(columns={
        "data documento":"data_nf",
        "descrição":"forma_pagamento",
        "documento":"os"
    },inplace=True)
    df["vencimento"] = pd.to_datetime(df["vencimento"],errors="coerce")
    # limpa registros 1970-01-01
    df.loc[df["vencimento"]==pd.Timestamp(1970,1,1),"vencimento"]=pd.NaT
    df["valor"]=pd.to_numeric(df["valor"],errors="coerce")
    df = df.dropna(subset=["fornecedor","valor"]).reset_index(drop=True)
    # calcula status_pagamento
    hoje = datetime.now().date()
    status = []
    for _,r in df.iterrows():
        e = str(r.get("estado","")).strip().lower()
        if e=="pago" or e=="recebido":
            status.append("Pago")
        else:
            d = r["vencimento"].date() if pd.notna(r["vencimento"]) else None
            if d:
                status.append("Em Atraso" if d<hoje else "A Vencer")
            else:
                status.append("Sem Data")
    df["status_pagamento"] = status
    return df

def save_data(path,month,df):
    wb=load_workbook(path); ws=wb[month]
    for i,r in df.iterrows():
        row = i+9
        ws.cell(row,5,r["vencimento"])
        ws.cell(row,6,r["valor"])
        ws.cell(row,7,r["estado"])
        ws.cell(row,8,r["situacao"])
    wb.save(path)

def delete_record(path,month,orig_idx):
    wb=load_workbook(path); ws=wb[month]
    # orig_idx é 0-based, header real começa na linha 9
    ws.delete_rows(orig_idx+9)
    wb.save(path)

def add_record(path,month,rec):
    wb=load_workbook(path); names=[s.strip() for s in wb.sheetnames]
    if month not in names:
        numeric=[s for s in names if s.isdigit()]
        template = wb[numeric[0] if numeric else names[0]]
        new=wb.copy_worksheet(template)
        new.title=month
        ws=new
    else:
        ws=wb[month]
    nxt=ws.max_row+1
    vals=[
        rec.get("data_nf",""),rec.get("forma_pagamento",""),rec.get("fornecedor",""),
        rec.get("os",""),rec.get("vencimento",""),rec.get("valor",""),
        rec.get("estado",""),rec.get("situacao",""),
        rec.get("boleto",""),rec.get("comprovante","")
    ]
    for idx,val in enumerate(vals, start=1):
        ws.cell(nxt,idx,val)
    wb.save(path)

# INTERFACE
st.sidebar.markdown("## 📂 Navegação")
page = st.sidebar.radio("",["Dashboard","Contas a Pagar","Contas a Receber"],index=0)

st.markdown("""
<div style="text-align:center;color:#4B8BBE;">
  <h1>💼 Sistema Financeiro 2025</h1>
  <p style="color:#555;">Dashboard avançado com estatísticas e gráficos interativos</p>
</div>
""",unsafe_allow_html=True)
st.markdown("---")

if page=="Dashboard":
    tab1,tab2 = st.tabs(["📥 Contas a Pagar","📤 Contas a Receber"])
    with tab1:
        sheets = get_existing_sheets(EXCEL_PAGAR)
        if not sheets: st.warning("Sem abas em Contas a Pagar")
        else:
            df = pd.concat([load_data(EXCEL_PAGAR,m) for m in sheets],ignore_index=True)
            st.metric("Total a Pagar",f"R$ {df['valor'].sum():,.2f}")
    with tab2:
        sheets = get_existing_sheets(EXCEL_RECEBER)
        if not sheets: st.warning("Sem abas em Contas a Receber")
        else:
            df = pd.concat([load_data(EXCEL_RECEBER,m) for m in sheets],ignore_index=True)
            st.metric("Total a Receber",f"R$ {df['valor'].sum():,.2f}")

elif page=="Contas a Pagar":
    st.subheader("🗂️ Contas a Pagar")
    aba = st.selectbox("Selecione o mês:",FULL_MONTHS)
    df = load_data(EXCEL_PAGAR,aba)
    view = st.radio("Visualizar:",["Todos","Pagas","Pendentes"],horizontal=True)
    if view=="Pagas":     df=df[df["estado"].str.lower()=="pago"]
    if view=="Pendentes": df=df[df["estado"].str.lower()!="pago"]
    disp = st.empty(); disp.dataframe(df,height=250)

    with st.expander("🗑️ Remover Registro"):
        if not df.empty:
            idx=st.number_input("Índice p/ remover:",0,len(df)-1,0,key="del_p")
            if st.button("Remover",key="btn_del_p"):
                delete_record(EXCEL_PAGAR,aba,df.index[idx])
                st.success("Registro removido!")
                df=load_data(EXCEL_PAGAR,aba)
                if view=="Pagas":     df=df[df["estado"].str.lower()=="pago"]
                if view=="Pendentes": df=df[df["estado"].str.lower()!="pago"]
                disp.dataframe(df,height=250)

elif page=="Contas a Receber":
    st.subheader("🗂️ Contas a Receber")
    aba = st.selectbox("Selecione o mês:",FULL_MONTHS)
    df = load_data(EXCEL_RECEBER,aba)
    view = st.radio("Visualizar:",["Todos","Recebidas","Pendentes"],horizontal=True)
    if view=="Recebidas": df=df[df["estado"].str.lower()=="recebido"]
    if view=="Pendentes": df=df[df["estado"].str.lower()!="recebido"]
    disp = st.empty(); disp.dataframe(df,height=250)

    with st.expander("🗑️ Remover Registro"):
        if not df.empty:
            idx=st.number_input("Índice p/ remover:",0,len(df)-1,0,key="del_r")
            if st.button("Remover",key="btn_del_r"):
                delete_record(EXCEL_RECEBER,aba,df.index[idx])
                st.success("Registro removido!")
                df=load_data(EXCEL_RECEBER,aba)
                if view=="Recebidas": df=df[df["estado"].str.lower()=="recebido"]
                if view=="Pendentes": df=df[df["estado"].str.lower()!="recebido"]
                disp.dataframe(df,height=250)

st.markdown("""
<div style="text-align:center;font-size:12px;color:gray;margin-top:20px;">
  © 2025 Desenvolvido por Vinicius Magalhães
</div>
""",unsafe_allow_html=True)
