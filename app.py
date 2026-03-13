import streamlit as st
import gspread
from google.oauth2 import service_account 
import os
import base64
import pandas as pd
from datetime import datetime
import time
import smtplib
import re
from email.message import EmailMessage
import plotly.express as px

# --- python -m streamlit run app.py ---
# --- CONFIGURAÇÕES E CAMINHOS ---
ID_PLANILHA = "1NS9zdzNFcHjQ7zFpEysuU-udrrV1VaM7nPY7LjHk3Qk"
ABA_USUARIOS = "SISAFA-NAVAL-Usuarios"
ABA_PROCESSOS = "SISAFA-NAVAL-processos"
ABA_LOGS_ACOES = "SISAFA-NAVAL-logs_acoes"
ABA_HISTORICO = "SISAFA-NAVAL-historico"
ABA_TABELA_A = "SISAFA-NAVAL-Tabela-A"

# Localiza a pasta do projeto
pasta_projeto = os.path.dirname(os.path.abspath(__file__))
caminho_logo = os.path.join(pasta_projeto, "LOGO-SISAFA-NAVAL.png")
caminho_mascote = os.path.join(pasta_projeto, "canto_inferior_direito_da_tela_de_apresentacao.png")

st.set_page_config(page_title="SISAFA-NAVAL (HNBra)", layout="centered", page_icon="⚓")

# --- ESTILIZAÇÃO CSS ---
st.markdown("""
    <style>
    [data-testid="stSidebarNav"] {display: none;} 
    .welcome-box { background: #f0f2f6; padding: 15px; border-radius: 10px; border-left: 8px solid #2e6b54; margin-bottom: 25px; font-size: 18px; font-weight: bold; color: #1B3129; }
    .stButton>button { background-color: #2e6b54; color: white; border-radius: 5px; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# --- FUNÇÕES DE CONEXÃO ---
def conectar_google():
    try:
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        if "gcp_service_account" in st.secrets:
            creds_info = st.secrets["gcp_service_account"]
            if isinstance(creds_info, str):
                import json
                creds_info = json.loads(creds_info.strip())
            
            # Garante que a chave privada seja lida corretamente em qualquer servidor
            if "private_key" in creds_info:
                creds_info["private_key"] = creds_info["private_key"].replace("\\n", "\n").strip()
            
            creds = service_account.Credentials.from_service_account_info(creds_info, scopes=scope)
            return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Erro na conexão com Google: {e}")
        return None

def registrar_historico(nup, fatura, origem, destino, valor, obs=""):
    try:
        client = conectar_google()
        sh = client.open_by_key(ID_PLANILHA)
        aba = sh.worksheet(ABA_HISTORICO)
        agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        aba.append_row([agora, nup, fatura, origem, destino, st.session_state.user_id, valor, obs])
    except Exception as e:
        st.error(f"Erro na aba HISTÓRICO: {e}")

def registrar_acao(nup, fatura, acao, detalhes=""):
    try:
        client = conectar_google()
        sh = client.open_by_key(ID_PLANILHA)
        aba = sh.worksheet(ABA_LOGS_ACOES)
        agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        aba.append_row([str(datetime.now().timestamp()), nup, fatura, acao, st.session_state.user_id, agora, detalhes])
    except: pass

def mover_status(nup, novo_status, auditor_nip=None, obs_texto=None, valor_glosa=None, valor_liq=None):
    client = conectar_google()
    sh = client.open_by_key(ID_PLANILHA)
    aba_p = sh.worksheet(ABA_PROCESSOS)
    cell = aba_p.find(nup)
    if cell:
        dados_atuais = aba_p.row_values(cell.row)
        status_origem = dados_atuais[10] 
        fatura = dados_atuais[4]
        valor_atual = valor_liq if valor_liq is not None else dados_atuais[7]
        
        agora = datetime.now().strftime("%d/%m/%Y %H:%M")
        aba_p.update_cell(cell.row, 11, novo_status)
        aba_p.update_cell(cell.row, 14, agora)
        if auditor_nip: aba_p.update_cell(cell.row, 12, auditor_nip)
        if obs_texto: aba_p.update_cell(cell.row, 18, obs_texto)
        if valor_glosa is not None: aba_p.update_cell(cell.row, 7, valor_glosa)
        if valor_liq is not None: aba_p.update_cell(cell.row, 8, valor_liq)
        
        registrar_historico(nup, fatura, status_origem, novo_status, valor_atual, obs_texto or "")
        return True
    return False

def disparar_email_glosa(destinatario, num_fatura, valor_glosa, justificativa, nome_ose, email_auditor):
    SMTP_SERVER = "smtp.gmail.com"
    SMTP_PORT = 587
    SMTP_USER = "hnbra.execucaofinanceira@gmail.com"
    
    # --- MUDANÇA DE SEGURANÇA AQUI ---
    # Agora ele busca a senha no "cofre" (Secrets) e não no texto plano
    SMTP_PASS = st.secrets["smtp_password"] 

    msg = EmailMessage()
    msg['Subject'] = f"Notificação de Glosa: Fatura {num_fatura} - HNBra"
    msg['From'] = SMTP_USER
    msg['To'] = destinatario
    
    if email_auditor:
        msg['Cc'] = email_auditor

    corpo_html = f"""
    <html>
        <body style="font-family: Arial, sans-serif;">
            <p>À (ao) <b>{nome_ose}</b>,</p>
            <p>Informamos que a auditagem da <b>Fatura nº {num_fatura}</b> resultou em uma glosa de <b>R$ {valor_glosa:,.2f}</b>.</p>
            <p><b>Justificativa:</b> {justificativa}</p>
            <p>O relatório de glosa seguirá formalmente assim que possível.</p>
            <br>
            <p>Cordialmente,</p>
            <p><b>Equipe de Auditoria em Saúde</b><br>
            Sistema de Acompanhamento de Faturas do Hospital Naval de Brasília</p>
            <hr>
            <p><small style="color: gray;">E-mail automático gerado pelo SISAFA-NAVAL. Favor não responder.</small></p>
        </body>
    </html>
    """
    msg.set_content("Favor visualizar em modo HTML.")
    msg.add_alternative(corpo_html, subtype='html')

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.send_message(msg)
        return True
    except Exception as e:
        st.error(f"Erro no envio: {e}")
        return False    

def limpar_valor(valor):
    if isinstance(valor, (int, float)): 
        return float(valor)
    limpo = str(valor).replace('R$', '').replace(' ', '').strip()
    if ',' in limpo and '.' in limpo:
        if limpo.find('.') < limpo.find(','): 
            limpo = limpo.replace('.', '').replace(',', '.')
        else: 
            limpo = limpo.replace(',', '')
    elif ',' in limpo:
        limpo = limpo.replace(',', '.')
    try:
        return float(limpo)
    except ValueError:
        return 0.0

# --- CONFIGURAÇÕES DE IMAGEM SEGURAS ---
pasta_projeto = os.path.dirname(os.path.abspath(__file__))
caminho_logo = os.path.join(pasta_projeto, "LOGO-SISAFA-NAVAL.png")
caminho_mascote = os.path.join(pasta_projeto, "canto_inferior_direito_da_tela_de_apresentacao.png")

# Função para carregar imagem sem quebrar o app
def carregar_imagem(caminho):
    return caminho if os.path.exists(caminho) else None


# --- CONEXÃO GLOBAL (Cole logo após a função limpar_valor) ---

# 1. Conecta ao Google e abre a planilha mestra
# --- INICIALIZAÇÃO DA CONEXÃO E CARREGAMENTO ---
client = conectar_google()

if client:
    try:
        # 1. Abre a planilha mestra
        sh = client.open_by_key(ID_PLANILHA)
        
        # 2. Define as abas globalmente
        aba_p = sh.worksheet(ABA_PROCESSOS)
        aba_l = sh.worksheet(ABA_LOGS_ACOES)
        aba_u = sh.worksheet(ABA_USUARIOS)
        aba_h = sh.worksheet(ABA_HISTORICO)
        
        # 3. Carrega os dados para o DataFrame principal
        dados = aba_p.get_all_records()
        df = pd.DataFrame(dados)
        
    except Exception as e:
        st.error(f"Conectado ao Google, mas erro ao carregar abas: {e}")
        df = pd.DataFrame() 
else:
    st.error("Não foi possível estabelecer a conexão inicial com o Google Sheets.")
    df = pd.DataFrame()

# --- CONTROLE DE SESSÃO ---
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'modulo_ativo' not in st.session_state: st.session_state.modulo_ativo = None
if 'confirmar_secom' not in st.session_state: st.session_state.confirmar_secom = False
if 'confirmar_recebimento' not in st.session_state: st.session_state.confirmar_recebimento = False
if 'confirmar_finalizacao' not in st.session_state: st.session_state.confirmar_finalizacao = False

# --- 1. TELA DE LOGIN ---
if not st.session_state.logged_in:
    # 1. Mascote (Mantido fixo no canto da tela, independente das colunas)
    mascote_path = carregar_imagem(caminho_mascote)
    if mascote_path:
        with open(mascote_path, "rb") as f:
            data = base64.b64encode(f.read()).decode()
            st.markdown(f'<img src="data:image/png;base64,{data}" style="position: fixed; bottom: 20px; right: 20px; width: 180px; z-index:999;">', unsafe_allow_html=True)

    # 2. Estrutura de Colunas para Centralização
    col1, col2, col3 = st.columns([1, 1.5, 1])
    
    with col2:
        # --- LOGO CENTRALIZADA ---
        # Ao colocar aqui dentro, ela segue o alinhamento da coluna central
        logo_path = carregar_imagem(caminho_logo)
        if logo_path: 
            # use_container_width garante que ela se ajuste ao espaço da coluna
            st.image(logo_path, use_container_width=True)
        else:
            st.markdown("<h1 style='text-align: center;'>⚓ SISAFA-NAVAL</h1>", unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True) # Pequeno espaço entre logo e campos

        # --- FORMULÁRIO ---
        tipo_acesso = st.radio("Tipo de Acesso:", ["Interno (NIP)", "Externo (CNPJ)"], horizontal=True)
        u_id = st.text_input(f"Digite seu {'NIP' if 'Interno' in tipo_acesso else 'CNPJ'}")
        senha = st.text_input("Senha", type="password")
        
        if st.button("ACESSAR SISTEMA", use_container_width=True):
            if client:
                try:
                    aba_u = sh.worksheet(ABA_USUARIOS)
                    usuarios = aba_u.get_all_values()
                    encontrado = False
                    for r in usuarios:
                        if str(r[0]).strip() == u_id.strip():
                            st.session_state.logged_in = True
                            st.session_state.user_id = u_id
                            st.session_state.user_full_name = r[1].upper()
                            st.session_state.user_perfil = r[2].upper()
                            encontrado = True
                            st.rerun()
                    if not encontrado: 
                        st.error("Usuário não cadastrado.")
                except Exception as e:
                    st.error(f"Erro ao validar login: {e}")

# --- 2. TELA DE SELEÇÃO DE MÓDULO (SALA DE ESPERA) ---
elif st.session_state.modulo_ativo is None:
    col_l1, col_l2, col_l3 = st.columns([1.2, 1, 1.2])
    with col_l2:
        if os.path.exists(caminho_logo): st.image(caminho_logo, use_container_width=True)
    
    st.markdown(f"<h1 style='text-align: center; color: #2e6b54;'>⚓ Bem-vindo, {st.session_state.user_full_name}</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 20px;'>Selecione o setor de trabalho abaixo para iniciar:</p><br>", unsafe_allow_html=True)
    
    perfis = [p.strip().upper() for p in st.session_state.user_perfil.split(',')]
    cols = st.columns(len(perfis))
    for i, perfil in enumerate(perfis):
        with cols[i]:
            if st.button(perfil, key=f"btn_{perfil}", use_container_width=True):
                st.session_state.modulo_ativo = perfil
                st.rerun()

# --- 3. AMBIENTE DE TRABALHO (MÓDULO ATIVO) ---
else:
    with st.sidebar:
        if os.path.exists(caminho_logo): st.image(caminho_logo)
        st.markdown(f"<p style='text-align:center;'><b>ID: {st.session_state.user_id}</b><br>Módulo: {st.session_state.modulo_ativo}</p>", unsafe_allow_html=True)
        if st.button("🔄 Trocar de Setor"):
            st.session_state.modulo_ativo = None
            st.rerun()
        if st.button("❌ Terminar Sessão"):
            st.session_state.logged_in = False
            st.session_state.modulo_ativo = None
            st.rerun()

    st.markdown(f'<div class="welcome-box">⚓ SISAFA-NAVAL: {st.session_state.modulo_ativo}</div>', unsafe_allow_html=True)
    
    client = conectar_google()
    sh = client.open_by_key(ID_PLANILHA)
    aba_p = sh.worksheet(ABA_PROCESSOS)
    df = pd.DataFrame(aba_p.get_all_records())

    # --- MÓDULOS ESPECÍFICOS ---

   
    if st.session_state.modulo_ativo == "SECOM" or st.session_state.modulo_ativo == "ADMIN":
        st.header("📥 Cadastro de Faturas (SECOM)")
        aba_a = sh.worksheet(ABA_TABELA_A)
        oses = {r[0].strip(): r[1].strip() for r in aba_a.get_all_values()[1:] if r[0]}
        
        nup_in = st.text_input("NUP (Ex: 63060.000123/2026-10)")
        
        c1, c2 = st.columns(2)
        sel_cnpj = c1.selectbox("Selecione o CNPJ da OSE", [""] + sorted(list(oses.keys())))
        empresa_nome = oses.get(sel_cnpj, "")
        c2.text_input("Empresa (OSE)", value=empresa_nome, disabled=True)
        
        num_fatura = st.text_input("Número da Fatura (Alfanumérico ou S/N)")
        v_ap = st.number_input("Valor Apresentado (R$)", min_value=0.0, format="%.2f")

        # --- BOTÃO DE PRÉ-CADASTRO ---
        if st.button("CADASTRAR FATURA"):
            if nup_in and sel_cnpj and num_fatura and v_ap > 0:
                st.session_state.confirmar_secom = True
            else:
                st.warning("⚠️ Preencha todos os campos obrigatórios antes de cadastrar.")

        # --- CAIXA DE CONFIRMAÇÃO ---
        if st.session_state.confirmar_secom:
            st.markdown("---")
            st.warning(f"**⚠️ CONFIRMAÇÃO:** Tem certeza de que os dados da fatura **{num_fatura}** estão corretos?")
            col_sim, col_nao = st.columns(2)
            
            if col_sim.button("✅ SIM, confirmar dados"):
                dt_hoje = datetime.now().strftime("%d/%m/%Y")
                
                # 1. Alimenta aba PROCESSOS (Snapshot)
                nova_linha = [
                    str(datetime.now().timestamp()), nup_in, sel_cnpj, empresa_nome, 
                    num_fatura, v_ap, 0, v_ap, datetime.now().month, datetime.now().year, 
                    1, st.session_state.user_id, dt_hoje, dt_hoje, "", "", "", ""
                ]
                aba_p.append_row(nova_linha)
                
                # 2. Alimenta aba HISTORICO (Macro - PM4PY)
                # IMPORTANTE: Verifique se na sua planilha o nome é 'Historico' ou 'historico'
                registrar_historico(nup_in, num_fatura, "0", "1", v_ap, "Entrada via SECOM")
                
                # 3. Alimenta aba LOGS_ACOES (Micro - Produtividade)
                registrar_acao(nup_in, num_fatura, "CADASTRO_INICIAL", f"Fatura cadastrada por {st.session_state.user_full_name}")
                
                # Feedback Visual
                st.success(f"🎉 Sucesso! Fatura {num_fatura} inserida no sistema.")
                st.session_state.confirmar_secom = False # Reseta a confirmação
                
                # Pequena pausa para o usuário ver o aviso antes de recarregar
                import time
                time.sleep(2)
                st.rerun()

            if col_nao.button("❌ NÃO, voltar e corrigir"):
                st.session_state.confirmar_secom = False
                st.rerun()
                







    elif st.session_state.modulo_ativo == "AUDITORIA" or st.session_state.modulo_ativo == "ADMIN":
        st.header("⚖️ Divisão de Auditoria em Saúde ⚕️")
        
        # Criação das 6 abas solicitadas
        t_fila, t_mesa, t_auditadas, t_busca, t_stats, t_rel = st.tabs([
            "📥 Fila de Espera", "🩺 Em Auditagem", "✅ Auditadas", 
            "🔍 Consultas", "📊 Produtividade", "💬 Relacionamento"
        ])

        # --- NOVO: MAPEAMENTO DE MESES PARA EXIBIÇÃO ---
        mapa_meses = {
        1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR", 5: "MAI", 6: "JUN",
        7: "JUL", 8: "AGO", 9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ"
        }
    
        # Criamos uma coluna nova 'mes_sigla' para visualização sem estragar os cálculos
        # Certificamos que 'mes_competencia' é numérico para o mapeamento funcionar
        df['mes_sigla'] = pd.to_numeric(df['mes_competencia'], errors='coerce').map(mapa_meses)

        # 1. ABA: FILA DE ESPERA
        with t_fila:
            # --- CÁLCULO DOS INDICADORES ---
            df_fila = df[df['status'] == 1].copy()
            
            if not df_fila.empty:
                # 1. Preparação dos dados: Limpeza de valores e conversão de datas
                df_fila['valor_limpo'] = df_fila['valor_apresentado'].apply(limpar_valor)
                df_fila['dt_entrada'] = pd.to_datetime(df_fila['data_entrada'], dayfirst=True, errors='coerce')
                
                hoje = datetime.now()
                df_fila['dias_fila'] = (hoje - df_fila['dt_entrada']).dt.days

            
                # 2. Cálculos de Temporalidade
                aceitavel = len(df_fila[df_fila['dias_fila'] <= 1])
                atencao = len(df_fila[(df_fila['dias_fila'] >= 2) & (df_fila['dias_fila'] <= 4)])
                atraso = len(df_fila[df_fila['dias_fila'] > 5])
                
                # 3. Valor Total na Fila
                valor_total_fila = df_fila['valor_limpo'].sum()

                # --- INTERFACE DE INDICADORES (KPIs) ---
                st.markdown("### 📊 Situação da fila de faturas cadastradas pela SECOM")
                
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Total de Faturas", f"{len(df_fila)}")
                c2.metric("🟢 Aceitável (até 1d)", f"{aceitavel}")
                c3.metric("🟡 Atenção (2-4d)", f"{atencao}")
                c4.metric("🔴 Em Atraso (>5d)", f"{atraso}")

                # Exibição do Valor por Competência
                with st.expander("💰 Detalhamento por Competência (Mês/Ano)", expanded=False):
                    resumo_comp = df_fila.groupby(['mes_sigla', 'ano_competencia'])['valor_limpo'].sum().reset_index()
                    resumo_comp.columns = ['Mês', 'Ano', 'Total Apresentado (R$)']
                    st.table(resumo_comp.style.format({'Total Apresentado (R$)': 'R$ {:,.2f}'}))
                
                st.divider()

            # --- TABELA DE PROCESSOS ---
            st.subheader("📥 Processos aguardando Auditoria")
            
            if df_fila.empty:
                st.info("Não há faturas na fila no momento.")
            else:
                # Adicionamos a coluna de Dias na Fila para facilitar a visão do auditor
                st.dataframe(
                    df_fila[['nup', 'ose', 'valor_apresentado', 'mes_sigla', 'ano_competencia', 'dias_fila']], 
                    use_container_width=True,
                    column_config={
                        "dias_fila": st.column_config.NumberColumn("Dias na Fila", help="Dias desde a entrada na SECOM")
                    }
                )
                
                # Seleção múltipla para recebimento em lote
                nups_selecionados = st.multiselect("Selecione os NUPs para trazer para sua mesa:", df_fila['nup'].tolist())
                
                if st.button("📥 RECEBER FATURA(S)"):
                    if nups_selecionados:
                        st.session_state.confirmar_recebimento = True
                    else:
                        st.warning("⚠️ Selecione ao menos uma fatura para receber.")

                # --- INTERFACE DE CONFIRMAÇÃO ---
                if st.session_state.confirmar_recebimento:
                    st.markdown("---")
                    st.warning(f"⚖️ **CONFIRMAÇÃO:** Você está prestes a assumir a responsabilidade por **{len(nups_selecionados)}** processo(s). Confirmar recebimento?")
                    
                    col_sim, col_nao = st.columns(2)
                    
                    if col_sim.button("✅ SIM, desejo receber"):
                        with st.spinner("Movimentando processos..."):
                            for n in nups_selecionados:
                                mover_status(n, 2, auditor_nip=st.session_state.user_id)
                                fatura_n = df[df['nup'] == n]['Numero_da_fatura'].values[0]
                                registrar_acao(n, fatura_n, "RECEBIMENTO_AUDITORIA", f"Auditor {st.session_state.user_id} puxou para a mesa.")

                        st.toast(f"✅ {len(nups_selecionados)} processos movidos!", icon="⚓")
                        st.session_state.confirmar_recebimento = False
                        time.sleep(1)
                        st.rerun()

                    if col_nao.button("❌ NÃO, cancelar"):
                        st.session_state.confirmar_recebimento = False
                        st.rerun()

        # 2. ABA: EM AUDITAGEM
        with t_mesa:
            # --- CÁLCULO DOS INDICADORES TÉCNICOS (Status 2) ---
            # Filtramos todos os processos em auditagem no sistema para os KPIs
            df_total_auditagem = df[df['status'] == 2].copy()
            
            if not df_total_auditagem.empty:
                # 1. Preparação dos dados
                df_total_auditagem['valor_limpo'] = df_total_auditagem['valor_apresentado'].apply(limpar_valor)
                
                # Usamos a coluna 14 (índice 13) que registra a data da entrada na auditoria
                # Caso sua planilha tenha um nome de cabeçalho específico, substitua .iloc[:, 13]
                df_total_auditagem['dt_mov'] = pd.to_datetime(df_total_auditagem.iloc[:, 13], dayfirst=True, errors='coerce')
                
                hoje = datetime.now()
                df_total_auditagem['dias_auditoria'] = (hoje - df_total_auditagem['dt_mov']).dt.days
                
                # 2. Cálculos de Temporalidade (Prazos da Auditoria)
                aceitavel_aud = len(df_total_auditagem[df_total_auditagem['dias_auditoria'] <= 10])
                atencao_aud = len(df_total_auditagem[(df_total_auditagem['dias_auditoria'] > 10) & (df_total_auditagem['dias_auditoria'] <= 15)])
                atraso_aud = len(df_total_auditagem[df_total_auditagem['dias_auditoria'] > 15])
                
                # --- INTERFACE DE INDICADORES (KPIs) ---
                st.markdown("### 📊 Situação Geral das Faturas em Auditagem")
                st.write("⚠️ **O número de dias é contado a partir do recebimento da fatura na Divisão de Auditoria** ⚠️")
                
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Total em Auditagem", f"{len(df_total_auditagem)}")
                c2.metric("🟢 Aceitável (até 10d)", f"{aceitavel_aud}")
                c3.metric("🟡 Atenção (11-15d)", f"{atencao_aud}")
                c4.metric("🔴 Em Atraso (>15d)", f"{atraso_aud}")

                # Detalhamento Financeiro em Status 2
                with st.expander("💰 Detalhamento por Competência (Mês/Ano)", expanded=False):
                    resumo_comp_aud = df_total_auditagem.groupby(['mes_sigla', 'ano_competencia'])['valor_limpo'].sum().reset_index()
                    resumo_comp_aud.columns = ['Mês', 'Ano', 'Total em Análise (R$)']
                    st.table(resumo_comp_aud.style.format({'Total em Análise (R$)': 'R$ {:,.2f}'}))
                
                st.divider()

            # --- MINHA MESA DE TRABALHO (VISÃO GERAL DO SETOR) ---
            st.subheader("🩺 Mesa de Trabalho da Auditoria")
            
            df['status'] = pd.to_numeric(df['status'], errors='coerce')
            
            # REMOVIDA A RESTRIÇÃO POR USER_ID: Agora filtra apenas pelo Status 2
            df_mesa = df[df['status'] == 2].copy()

            if df_mesa.empty:
                st.info("Não há processos em auditagem no momento.")
            else:
                st.write("**Todas as faturas em análise técnica:**")
                # Incluímos a coluna de dias em auditoria para gestão
                if not df_total_auditagem.empty:
                    # Cálculo dos dias baseado na coluna 14 (índice 13)
                    df_mesa['dias_auditoria'] = (hoje - pd.to_datetime(df_mesa.iloc[:, 13], dayfirst=True, errors='coerce')).dt.days
  
                st.dataframe(df_mesa[['nup', 'ose', 'valor_apresentado', 'mes_sigla', 'ano_competencia', 'obs']], use_container_width=True)
                
                st.divider()
                
                nup_audit = st.selectbox(
                    "Selecione o NUP para realizar a análise técnica:", 
                    [""] + df_mesa['nup'].tolist(),
                    key="sb_nup_analise_mesa_final"
                )
                
                if nup_audit:
                    dados_nup = df_mesa[df_mesa['nup'] == nup_audit].iloc[0]
                    num_fat = dados_nup['Numero_da_fatura']
                    v_apres = limpar_valor(dados_nup['valor_apresentado'])
                    
                    st.markdown(f"#### 📝 Analisando Fatura: **{num_fat}**")
                    
                    c1, c2 = st.columns(2)
                    with c1:
                        glosa_input = st.number_input("Valor da Glosa (R$)", min_value=0.0, max_value=v_apres, step=0.01, format="%.2f", key="val_glosa_mesa")
                        just_glosa = st.text_area("Justificativa Técnica da Glosa", height=150, key="txt_just_mesa")
                    
                    with c2:
                        v_liquido = v_apres - glosa_input
                        st.metric("Valor Apresentado", f"R$ {v_apres:,.2f}")
                        st.metric("Valor Líquido Final", f"R$ {v_liquido:,.2f}", delta=f"- R$ {glosa_input:,.2f}" if glosa_input > 0 else None, delta_color="inverse")

                    # --- PAINEL DE CONFERÊNCIA DE SEGURANÇA ---
                    st.markdown("### 🛡️ Conferência de Dados e Destino")
                    with st.container(border=True):
                        try:
                            # 1. Busca dados da OSE na Tabela-A
                            cnpj_fatura = str(dados_nup['cnpj']).strip().split('.')[0]
                            df_ose_info = pd.DataFrame(sh.worksheet(ABA_TABELA_A).get_all_records())
                            linha_ose = df_ose_info[df_ose_info['CNPJ'].astype(str).str.contains(cnpj_fatura)]
                            
                            # 2. Busca e-mail do Auditor logado
                            df_users = pd.DataFrame(sh.worksheet(ABA_USUARIOS).get_all_records())
                            match_user = df_users[df_users['NIP'].astype(str).str.strip() == user_logado]
                            
                            if not match_user.empty:
                                col_email = 'Email' if 'Email' in match_user.columns else 'E-mail'
                                email_auditor = match_user[col_email].values[0]
                            else:
                                email_auditor = None
                                st.error(f"⚠️ Seu NIP ({user_logado}) não foi encontrado na aba Usuários.")

                            if not linha_ose.empty and email_auditor:
                                email_destino = linha_ose['E-mail Principal da OSE'].values[0]
                                nome_ose_oficial = linha_ose['Razão Social'].values[0]
                                
                                st.write(f"🏢 **Organização de Saúde:** {nome_ose_oficial}")
                                st.write(f"📩 **E-mail de Destino:** {email_destino}")
                                st.write(f"📎 **Em Cópia (CC):** {email_auditor}")
                                
                                trava_confirmacao = st.checkbox("Confirmo que a OSE e os valores de glosa estão corretos.")
                            else:
                                if linha_ose.empty: st.error(f"⚠️ CNPJ {cnpj_fatura} não localizado na Tabela-A.")
                                trava_confirmacao = False
                        except Exception as e:
                            st.error(f"Erro ao carregar dados: {e}")
                            trava_confirmacao = False

                    st.divider()
                    
                    col_fin, col_mail = st.columns(2)
                    
                    if col_fin.button("✅ FINALIZAR AUDITORIA", use_container_width=True, key="btn_fin_mesa"):
                        if glosa_input > 0 and not just_glosa:
                            st.error("⚠️ Preencha a justificativa para aplicar a glosa.")
                        else:
                            st.session_state.confirmar_finalizacao = True

                    if col_mail.button("📧 ENCAMINHAR GLOSA P/ OSE", use_container_width=True, key="btn_mail_mesa", disabled=not trava_confirmacao):
                        with st.spinner("Enviando comunicado..."):
                            sucesso = disparar_email_glosa(
                                destinatario=email_destino,
                                num_fatura=num_fat,
                                valor_glosa=glosa_input,
                                justificativa=just_glosa,
                                nome_ose=nome_ose_oficial,
                                email_auditor=email_auditor
                            )
                            if sucesso:
                                registrar_acao(nup_audit, num_fat, "EMAIL_GLOSA_ENVIADO", f"Para: {email_destino} | CC: {email_auditor}")
                                st.toast(f"E-mail enviado com sucesso!", icon="✅")

                    # --- INTERFACE DE CONFIRMAÇÃO DE FINALIZAÇÃO ---
                    if st.session_state.confirmar_finalizacao:
                        st.markdown("---")
                        st.warning(f"🛡️ **CONFIRMAÇÃO:** Finalizar NUP **{nup_audit}** com Líquido de **R$ {v_liquido:,.2f}**?")
                        b_sim, b_nao = st.columns(2)
                        
                        if b_sim.button("👍 SIM, Finalizar", key="ok_fin_final"):
                            mover_status(nup_audit, 3, valor_glosa=glosa_input, valor_liq=v_liquido, obs_texto=just_glosa)
                            registrar_acao(nup_audit, num_fat, "AUDITORIA_CONCLUIDA", f"Glosa: {glosa_input}")
                            st.success("Auditoria concluída!")
                            st.session_state.confirmar_finalizacao = False
                            time.sleep(1.5)
                            st.rerun()

                        if b_nao.button("🔙 Voltar", key="no_fin_final"):
                            st.session_state.confirmar_finalizacao = False
                            st.rerun()


        # 3. ABA: FATURAS AUDITADAS
        with t_auditadas:
            # --- CÁLCULO DOS INDICADORES (Status 3) ---
            df_auditadas = df[df['status'] == 3].copy()
            
            if not df_auditadas.empty:
                # 1. Preparação dos dados
                df_auditadas['v_apres_limpo'] = df_auditadas['valor_apresentado'].apply(limpar_valor)
                df_auditadas['glosa_limpo'] = df_auditadas['glosa'].apply(limpar_valor)
                df_auditadas['v_liq_limpo'] = df_auditadas['valor_liquido'].apply(limpar_valor)
                
                # Regra de data: Coluna 14 (Índice 13)
                df_auditadas['dt_entrada_setor'] = pd.to_datetime(df_auditadas.iloc[:, 13], dayfirst=True, errors='coerce')
                
                hoje = datetime.now()
                df_auditadas['dias_no_setor'] = (hoje - df_auditadas['dt_entrada_setor']).dt.days
                
                # 2. Parâmetros de Temporalidade (10, 15 e >15)
                aceitavel_aud = len(df_auditadas[df_auditadas['dias_no_setor'] <= 10])
                atencao_aud = len(df_auditadas[(df_auditadas['dias_no_setor'] >= 11) & (df_auditadas['dias_no_setor'] <= 15)])
                atraso_aud = len(df_auditadas[df_auditadas['dias_no_setor'] > 15])

                # --- INTERFACE DE INDICADORES ---
                st.markdown("### 📊 Faturas Auditadas aguardando encaminhamento")
                st.write("⚠️ **O número de dias é contado a partir do recebimento da fatura na Divisão de Auditoria** ⚠️")
                
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Total Concluído", f"{len(df_auditadas)}")
                c2.metric("🟢 No Prazo (até 10d)", f"{aceitavel_aud}")
                c3.metric("🟡 Atenção (11-15d)", f"{atencao_aud}")
                c4.metric("🔴 Em Atraso (>15d)", f"{atraso_aud}")

                with st.expander("💰 Resumo Financeiro da Produção Geral", expanded=False):
                    total_ap = df_auditadas['v_apres_limpo'].sum()
                    total_gl = df_auditadas['glosa_limpo'].sum()
                    total_lq = df_auditadas['v_liq_limpo'].sum()
                    st.write(f"**Valor Apresentado Total:** R$ {total_ap:,.2f}")
                    st.write(f"**Glosas:** R$ {total_gl:,.2f}")
                    st.write(f"**Líquido a Pagar:** R$ {total_lq:,.2f}")
                
                st.divider()

                # --- LISTAGEM E AÇÕES ---
                st.subheader("✅ Faturas Prontas para Encaminhamento")
                st.dataframe(
                    df_auditadas[['nup', 'ose', 'valor_apresentado', 'glosa', 'valor_liquido', 'mes_sigla', 'ano_competencia', 'dias_no_setor']], 
                    use_container_width=True
                )
                
                lote_exec = st.multiselect(
                    "Selecionar faturas para encaminhar à Execução Financeira:", 
                    df_auditadas['nup'].tolist(), 
                    key="ms_lote_auditadas_v2"
                )
                
                if st.button("📤 ENCAMINHAR PARA EXECUÇÃO FINANCEIRA", key="btn_envio_fin_v2", use_container_width=True):
                    if lote_exec:
                        with st.spinner("Registrando encaminhamento..."):
                            for n in lote_exec:
                                fatura_n = df_auditadas[df_auditadas['nup'] == n]['Numero_da_fatura'].values[0]
                                registrar_acao(n, fatura_n, "ENCAMINHADO_PARA_FINANCEIRO", f"Usuário {st.session_state.user_id} encaminhou o lote.")
                        st.success(f"✅ {len(lote_exec)} processos notificados.")
                        time.sleep(1.5)
                        st.rerun()
            else:
                st.info("Nenhuma fatura auditada pendente no momento.")


        # 4. ABA: CONSULTAS (BUSCA GLOBAL)
        with t_busca:
            st.subheader("🔍 Localizar e Rastrear Processo")
            
            # Atualizamos o placeholder para incluir "Fatura"
            termo_busca = st.text_input(
                "Pesquise por NUP, Empresa (OSE), CNPJ ou Número da Fatura:", 
                placeholder="Ex: HFA, 63060, 123/2026..."
            )

            if termo_busca:
                # Filtragem flexível expandida para incluir a coluna 'Numero_da_fatura'
                mask = (
                    df['nup'].astype(str).str.contains(termo_busca, case=False, na=False) |
                    df['ose'].astype(str).str.contains(termo_busca, case=False, na=False) |
                    df['cnpj'].astype(str).str.contains(termo_busca, case=False, na=False) |
                    df['Numero_da_fatura'].astype(str).str.contains(termo_busca, case=False, na=False)
                )
                df_resultados = df[mask]

                if df_resultados.empty:
                    st.warning("Nenhum processo localizado com esse termo.")
                else:
                    st.write(f"📂 **{len(df_resultados)}** processo(s) encontrado(s):")
                    # Exibição resumida para conferência
                    st.dataframe(
                        df_resultados[['nup', 'ose', 'Numero_da_fatura', 'valor_apresentado', 'status']], 
                        use_container_width=True
                    )
                    
                    st.divider()
                    
                    # Seleção do NUP específico para ver o detalhamento completo
                    lista_nups = df_resultados['nup'].tolist()
                    nup_selecionado = st.selectbox("Selecione o NUP para ver a Linha do Tempo detalhada:", [""] + lista_nups)

                    if nup_selecionado:
                        # --- 1. SNAPSHOT ATUAL ---
                        res = df[df['nup'] == nup_selecionado].iloc[0]
                        
                        c1, c2, c3 = st.columns(3)
                        c1.metric("Status Atual", f"Fase {res['status']}")
                        c2.metric("Responsável", res['responsavel_atual'])
                        c3.metric("Valor Líquido", f"R$ {limpar_valor(res['valor_liquido']):,.2f}")

                        # --- 2. TRILHA DE AUDITORIA (HISTÓRICO MACRO) ---
                        st.markdown("### 👣 Trilha de Auditoria (Status)")
                        try:
                            aba_h = sh.worksheet(ABA_HISTORICO)
                            hist_df = pd.DataFrame(aba_h.get_all_records())
                            track = hist_df[hist_df['nup'] == nup_selecionado].sort_values(by='timestamp')
                            
                            if not track.empty:
                                for _, row in track.iterrows():
                                    st.caption(f"🕒 {row['timestamp']} | De: **{row['status_origem']}** ⮕ Para: **{row['status_destino']}** | Usuário: {row['usuario']}")
                            else:
                                st.info("Sem histórico de movimentação registrado.")
                        except:
                            st.error("Erro ao carregar histórico.")

                        # --- 3. AÇÕES E TRÂMITES (LOGS MICRO) ---
                        st.markdown("### 📧 Ações, E-mails e Trâmites")
                        try:
                            aba_l = sh.worksheet(ABA_LOGS_ACOES)
                            logs_df = pd.DataFrame(aba_l.get_all_records())
                            acoes = logs_df[logs_df['nup'] == nup_selecionado].sort_values(by='data_hora', ascending=False)
                            
                            if not acoes.empty:
                                st.table(acoes[['data_hora', 'acao', 'militar_nip', 'detalhes']])
                            else:
                                st.info("Nenhuma ação específica registrada para este NUP.")
                        except:
                            st.error("Erro ao carregar logs de ações.")



    # 5. ABA: PRODUTIVIDADE E ESTATÍSTICAS
        with t_stats:
            st.header("📈 Inteligência de Dados e Produtividade")

            # --- FILTROS DE COMPETÊNCIA ---
            col_f1, col_f2 = st.columns(2)
            anos_disp = sorted(df['ano_competencia'].unique(), reverse=True)
            ano_sel = col_f1.selectbox("Filtrar por Ano:", ["Todos"] + list(anos_disp))
            
            meses_disp = sorted(df['mes_sigla'].unique())
            mes_sel = col_f2.selectbox("Filtrar por Mês:", ["Todos"] + list(meses_disp))

            # Filtragem dinâmica
            df_p = df.copy()
            if ano_sel != "Todos": df_p = df_p[df_p['ano_competencia'] == ano_sel]
            if mes_sel != "Todos": df_p = df_p[df_p['mes_sigla'] == mes_sel]

            df_p['v_ap_num'] = df_p['valor_apresentado'].apply(limpar_valor)
            df_p['glosa_num'] = df_p['glosa'].apply(limpar_valor)

            # --- 1. SEÇÃO: PIZZAS E VOLUMES ---
            st.subheader("📌 Visão Geral do Volume")
            c1, c2, c3 = st.columns(3)
            c1.metric("Processos", len(df_p))
            c2.metric("Total Apresentado", f"R$ {df_p['v_ap_num'].sum():,.2f}")
            c3.metric("Total Glosado", f"R$ {df_p['glosa_num'].sum():,.2f}")

            cp1, cp2, cp3 = st.columns(3)
            # Pizza 1: Auditoria vs Auditadas
            st_counts = df_p[df_p['status'].isin([2, 3])]['status'].map({2:'Em Mesa', 3:'Concluídas'}).value_counts().reset_index()
            if not st_counts.empty:
                cp1.plotly_chart(px.pie(st_counts, values='count', names='status', title="Mesa vs. Concluídas", hole=0.4), use_container_width=True)
            
            # Pizza 2: Impacto Financeiro
            v_t = df_p['v_ap_num'].sum()
            v_g = df_p['glosa_num'].sum()
            if v_t > 0:
                cp2.plotly_chart(px.pie(values=[v_t - v_g, v_g], names=['Líquido', 'Glosa'], title="Impacto da Glosa", color_discrete_sequence=['#2e6b54', '#d32f2f']), use_container_width=True)
            
            # Pizza 3: Top 10 OSEs
            top_ose = df_p.groupby('ose')['v_ap_num'].sum().sort_values(ascending=False).head(10).reset_index()
            if not top_ose.empty:
                cp3.plotly_chart(px.pie(top_ose, values='v_ap_num', names='ose', title="Top 10 OSEs (Valor)"), use_container_width=True)

            # --- 2. SEÇÃO: TERMÔMETRO E SAÚDE ---
            st.divider()
            st.subheader("🌡️ Termômetro de Saúde do Processo (Global)")
            
            hoje = datetime.now()
            df_p['dt_ent'] = pd.to_datetime(df_p['data_entrada'], dayfirst=True, errors='coerce')
            df_p['dias_hoje'] = (hoje - df_p['dt_ent']).dt.days

            def classificar_global(d):
                if pd.isna(d): return "Desconhecido"
                if d <= 15: return "🟢 Aceitável"
                if d <= 25: return "🟡 Atenção"
                return "🔴 Em Atraso"

            df_p['situacao'] = df_p['dias_hoje'].apply(classificar_global)
            
            # Gráfico de barras de saúde
            saude_counts = df_p['situacao'].value_counts().reset_index()
            if not saude_counts.empty:
                fig_saude = px.bar(saude_counts, x='situacao', y='count', color='situacao', 
                                   title="Saúde do Passivo (Desde o Cadastro)",
                                   color_discrete_map={"🟢 Aceitável": "#2e6b54", "🟡 Atenção": "#f1c40f", "🔴 Em Atraso": "#e74c3c"})
                st.plotly_chart(fig_saude, use_container_width=True)

    # 6. ABA: RELACIONAMENTO
        with t_rel:
            st.subheader("💬 Central de Relacionamento (Inbox OSE)")
            
            try:
                # 1. Carregamos as mensagens da aba correspondente
                aba_msg = sh.worksheet(ABA_MENSAGENS)
                df_msg = pd.DataFrame(aba_msg.get_all_records())
                
                if df_msg.empty:
                    st.info("Nenhuma mensagem ou questionamento pendente no momento.")
                else:
                    # 2. Métricas Rápidas
                    pendentes = len(df_msg[df_msg['status_resposta'] == 'PENDENTE'])
                    media_reserva = "2.4 dias" # Placeholder para cálculo futuro
                    
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Total de Mensagens", len(df_msg))
                    c2.metric("📩 Pendentes", pendentes, delta=f"{pendentes} aguardando", delta_color="inverse")
                    c3.metric("⏱️ Tempo Médio de Resposta", media_reserva)
                    
                    st.divider()

                    # 3. Inbox de Mensagens
                    st.write("**📥 Mensagens Recebidas:**")
                    # Filtramos para mostrar primeiro o que não foi respondido
                    df_exibir = df_msg.sort_values(by='status_resposta', ascending=False)
                    st.dataframe(
                        df_exibir[['nup', 'cnpj_ose', 'assunto', 'data_envio', 'status_resposta']], 
                        use_container_width=True
                    )

                    # 4. Área de Resposta Técnica
                    st.markdown("### ✍️ Responder Questionamento")
                    nup_alvo = st.selectbox("Selecione o NUP para responder:", [""] + df_msg['nup'].unique().tolist())
                    
                    if nup_alvo:
                        msg_data = df_msg[df_msg['nup'] == nup_alvo].iloc[0]
                        
                        with st.container(border=True):
                            st.write(f"**De (OSE):** {msg_data['cnpj_ose']}")
                            st.write(f"**Assunto:** {msg_data['assunto']}")
                            st.info(f"**Mensagem da OSE:** {msg_data['mensagem_corpo']}")
                            
                            resposta_texto = st.text_area("Resposta Oficial do Auditor:", height=150, placeholder="Digite aqui o parecer técnico...")
                            
                            col_env, _ = st.columns([1, 2])
                            if col_env.button("📤 ENVIAR RESPOSTA OFICIAL", use_container_width=True):
                                if resposta_texto:
                                    # Lógica futura: Gravar resposta na planilha e disparar e-mail
                                    registrar_acao(nup_alvo, "N/A", "RESPOSTA_OSE", f"Auditor respondeu questionamento via sistema.")
                                    st.success("Resposta enviada com sucesso para o Portal da OSE!")
                                    time.sleep(1.5)
                                    st.rerun()
                                else:
                                    st.warning("Escreva uma resposta antes de enviar.")

            except Exception as e:
                st.error(f"Erro ao carregar Central de Relacionamento: {e}")    



        

    elif "EXECUÇÃO" in st.session_state.modulo_ativo or st.session_state.modulo_ativo == "ADMIN":
        st.header("💰 Execução Financeira")

        # Criando as abas
        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "📥 Fila de Espera", "📄 Gestão de NE", "💸 Gestão de Pagamentos", 
            "📊 Estatísticas e Indicadores", "🔍 Consultas", "🤝 Relacionamento"
        ])

        # --- ABA 1: FILA DE ESPERA (Baseada no Layout da Auditoria) ---
        with tab1:
            # --- CÁLCULO DOS INDICADORES ---
            df_fila_exec = df[df['status'] == 3].copy()
            
            # Cruzamento automático com Logs para pegar a data de encaminhamento
            logs_raw = aba_l.get_all_records()
            df_logs = pd.DataFrame(logs_raw)
            df_envio = df_logs[df_logs['acao'] == "ENCAMINHADO_PARA_FINANCEIRO"].copy()
            df_envio['dt_chegada'] = pd.to_datetime(df_envio['data_hora'], dayfirst=True, errors='coerce')
            df_envio = df_envio.sort_values('dt_chegada', ascending=False).drop_duplicates('nup')
            
            # Mescla a data do log no dataframe de execução
            df_fila_exec = df_fila_exec.merge(df_envio[['nup', 'dt_chegada']], on='nup', how='left')

            if not df_fila_exec.empty:
                # 1. Preparação dos dados: Limpeza de valores e cálculo de dias
                # (Usando o v_num que criamos para evitar erro de string)
                df_fila_exec['valor_limpo'] = df_fila_exec['valor_liquido'].astype(str).str.replace('R$', '', regex=False).str.replace('.', '', regex=False).str.replace(',', '.', regex=False).str.strip()
                df_fila_exec['valor_limpo'] = pd.to_numeric(df_fila_exec['valor_limpo'], errors='coerce').fillna(0.0)
                
                hoje = datetime.now()
                df_fila_exec['dias_fila'] = (hoje - df_fila_exec['dt_chegada']).dt.days.fillna(0).astype(int)

                # 2. Cálculos de Temporalidade (Padrão Execução: 2d, 3d, 4d+)
                aceitavel = len(df_fila_exec[df_fila_exec['dias_fila'] <= 2])
                atencao = len(df_fila_exec[df_fila_exec['dias_fila'] == 3])
                atraso = len(df_fila_exec[df_fila_exec['dias_fila'] >= 4])

                # --- INTERFACE DE INDICADORES (KPIs) ---
                st.markdown("### 📊 Situação da fila de faturas enviadas pela Auditoria")
                
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Total de Faturas", f"{len(df_fila_exec)}")
                c2.metric("🟢 Aceitável (até 2d)", f"{aceitavel}")
                c3.metric("🟡 Atenção (3d)", f"{atencao}")
                c4.metric("🔴 Em Atraso (>= 4d)", f"{atraso}")

                # Detalhamento por Competência
                with st.expander("💰 Detalhamento por Competência (Mês/Ano)", expanded=False):
                    resumo_comp = df_fila_exec.groupby(['mes_competencia','ano_competencia'])['valor_limpo'].sum().reset_index()
                    resumo_comp.columns = ['Mês','Ano','Total Líquido (R$)']
                    st.table(resumo_comp.style.format({'Total Líquido (R$)':'R${:,.2f}'}))
                
                st.divider()

            # --- TABELA DE PROCESSOS ---
            st.subheader("📥 Processos aguardando Recebimento no Financeiro")
            
            if df_fila_exec.empty:
                st.info("Não há faturas na fila no momento.")
            else:
                st.dataframe(
                    df_fila_exec[['nup','ose','valor_liquido','mes_competencia','ano_competencia','dias_fila']],
                    use_container_width=True,
                    column_config={
                        "dias_fila": st.column_config.NumberColumn("Dias na Fila", help="Dias desde o envio pela Auditoria")
                    }
                )
                
                # Seleção múltipla para recebimento
                nups_sel = st.multiselect("Selecione os NUPs para receber:", df_fila_exec['nup'].tolist(), key="multi_exec")
                
                if st.button("📥 RECEBER FATURA(S)"):
                    if nups_sel:
                        st.session_state.confirma_exec = True
                    else:
                        st.warning("⚠️ Selecione ao menos uma fatura.")

                # --- INTERFACE DE CONFIRMAÇÃO ---
                if st.session_state.get('confirma_exec', False):
                    st.markdown("---")
                    st.warning(f"⚖️ **CONFIRMAÇÃO:** Você vai receber **{len(nups_sel)}** processo(s). Confirmar?")
                    
                    col_sim, col_nao = st.columns(2)
                    
                    if col_sim.button("✅ SIM, receber"):
                        with st.spinner("Movimentando..."):
                            for n in nups_sel:
                                mover_status(n, 4) # Move para Aguardando emissão de NE
                                fatura_n = df[df['nup'] == n]['Numero_da_fatura'].values[0]
                                registrar_acao(n, fatura_n, "RECEBIMENTO_FINANCEIRO", f"Financeiro {st.session_state.user_id} recebeu a fatura.")
                        
                        st.success(f"✅ {len(nups_sel)} processos movidos!")
                        st.session_state.confirma_exec = False
                        time.sleep(1)
                        st.rerun()

                    if col_nao.button("❌ NÃO, cancelar"):
                        st.session_state.confirma_exec = False
                        st.rerun()

        # --- ABA 2: GESTÃO DE NE (Status 4 -> 5 -> 6) ---
        with tab2:
            st.markdown("### 📝 1. Emitir Nota de Empenho (NE)")
            
            # 1. Filtro de faturas prontas para empenho
            f_status_4 = df[df['status'] == 4].copy()
            
            if not f_status_4.empty:
                # Usamos a coluna 'mes_sigla' que criamos no início do módulo
                nups_sel = st.multiselect(
                    "Selecione o(s) NUP(s) para empenhar (Devem ser da mesma empresa):", 
                    f_status_4['nup'].tolist(), 
                    key="sel_ne_batch"
                )
                
                # --- BLOCO DE AJUDA E CONFERÊNCIA ---
                trava_cnpj = False
                if nups_sel:
                    df_conf = f_status_4[f_status_4['nup'].isin(nups_sel)].copy()
                    df_conf['v_liq_num'] = df_conf['valor_liquido'].apply(limpar_valor)
                    
                    # Verificação de CNPJs únicos
                    lista_cnpjs = df_conf['cnpj'].unique()
                    trava_cnpj = len(lista_cnpjs) > 1
                    
                    # Dados para o Card
                    empresa_nome = df_conf['ose'].iloc[0]
                    cnpj_principal = df_conf['cnpj'].iloc[0]
                    valor_total_ne = df_conf['v_liq_num'].sum()
                    faturas_no_lote = ", ".join(df_conf['Numero_da_fatura'].astype(str).tolist())

                    with st.container(border=True):
                        st.markdown(f"#### 🔎 Conferência de Empenho")
                        c_aj1, c_aj2 = st.columns([2, 1])
                        
                        with c_aj1:
                            st.write(f"🏢 **Empresa:** {empresa_nome}")
                            st.write(f"🆔 **CNPJ:** {cnpj_principal}")
                            st.write(f"📄 **Faturas:** {faturas_no_lote}")
                            
                            if trava_cnpj:
                                st.error("❌ **ALERTA:** Múltiplos CNPJs detectados. Remova os NUPs intrusos.")
                        
                        with c_aj2:
                            st.metric("Qtd. Faturas", len(df_conf))
                            st.metric("Total da NE", f"R$ {valor_total_ne:,.2f}")

                # --- INPUTS E CADASTRO ---
                col_input1, col_input2 = st.columns([1,1])
                with col_input1:
                    cod_ne_final = st.text_input("Número Final da NE (ex: 00052)", key="input_ne_num")
                
                # O botão fica desabilitado se houver erro de CNPJ ou nada selecionado
                if st.button("🚀 Cadastrar NE", disabled=trava_cnpj or not nups_sel, use_container_width=True):
                    if not cod_ne_final:
                        st.warning("⚠️ Digite o número da NE antes de prosseguir.")
                    else:
                        ne_completa = f"78770000001NE{datetime.now().year}{cod_ne_final}"
                        cnpj_alvo = f_status_4[f_status_4['nup'].isin(nups_sel)]['cnpj'].iloc[0]
                        
                        with st.spinner(f"Gravando NE {ne_completa}..."):
                            for nup in nups_sel:
                                cell = aba_p.find(nup)
                                if cell:
                                    # Grava a NE na Coluna O (15)
                                    aba_p.update_cell(cell.row, 15, ne_completa)
                                    # Move para Status 5 (Empenhado)
                                    mover_status(nup, 5)
                                    
                                    fatura_n = df[df['nup'] == nup]['Numero_da_fatura'].values[0]
                                    registrar_acao(nup, fatura_n, "NE_CADASTRADA", f"NE {ne_completa} vinculada ao CNPJ {cnpj_alvo}")
                            
                            st.success(f"✅ Sucesso! NE {ne_completa} cadastrada para {empresa_nome}.")
                            time.sleep(1.5)
                            st.rerun()

                st.divider()
                st.subheader("📋 Processos Disponíveis para Empenho")
                cols_v = ['nup','cnpj','ose','mes_sigla','ano_competencia','valor_liquido']
                st.dataframe(f_status_4[cols_v].rename(columns={'mes_sigla':'Mês'}), use_container_width=True)
            
            else:
                st.info("Não há faturas aguardando emissão de NE.")

            st.divider()

            # --- SEÇÃO 2: ENVIO PARA FISCALIZAÇÃO (Status 5 -> 6) ---
        st.markdown("### 📤 2. Encaminhar para Fiscalização")
        
        # Filtra apenas Status 5 (Já empenhadas)
        f_status_5 = df[df['status'] == 5].copy()
        
        if not f_status_5.empty:
            # Mapeamento de meses para visualização
            meses_siglas = {
                1:'JAN', 2:'FEV', 3:'MAR', 4:'ABR', 5:'MAI', 6:'JUN',
                7:'JUL', 8:'AGO', 9:'SET', 10:'OUT', 11:'NOV', 12:'DEZ'
            }
            f_status_5['mes_extenso'] = f_status_5['mes_competencia'].map(meses_siglas).fillna(f_status_5['mes_competencia'])

            # Seleção em lote para o fiscal
            lote_fiscal = st.multiselect(
                "Selecione o(s) NUP(s) para enviar ao Fiscal:", 
                f_status_5['nup'].tolist(), 
                key="lote_fiscal_exec"
            )

            if st.button("📧 Enviar em Lote"):
                if lote_fiscal:
                    with st.spinner("Movimentando processos para Fiscalização..."):
                        for n in lote_fiscal:
                            # Evolui para Status 6 (Em Fiscalização)
                            mover_status(n, 6)
                            
                            # Log da ação
                            fatura_n = df[df['nup'] == n]['Numero_da_fatura'].values[0]
                            registrar_acao(n, fatura_n, "ENVIO_FISCALIZACAO", "Encaminhado para conferência do Fiscal (PJS/Fiscais).")
                    
                    st.success(f"✅ {len(lote_fiscal)} faturas encaminhadas para fiscalização com sucesso!")
                    time.sleep(1)
                    st.rerun()
                else:
                    st.warning("⚠️ Selecione ao menos um NUP da lista abaixo.")

            # --- TABELA ORGANIZADA POR EMPENHO ---
            st.subheader("Processos Empenhados (Aguardando Envio)")
            
            # Colunas solicitadas: NUP, CNPJ, OSE, Valor, Mês, Ano e NE
            cols_v = ['nup', 'cnpj', 'ose', 'valor_liquido', 'mes_extenso', 'ano_competencia', 'ne']
            
            # Ordenamos pela coluna 'ne' para agrupar visualmente as faturas do mesmo empenho
            df_exibir_fiscal = f_status_5[cols_v].sort_values(by='ne').rename(columns={'mes_extenso': 'Mês', 'ne': 'Nota de Empenho'})
            
            st.dataframe(df_exibir_fiscal, use_container_width=True)
            
        else:
            st.info("Nenhuma fatura empenhada (Status 5) aguardando envio.")

        # --- ABA 3: GESTÃO DE PAGAMENTOS (Status 7 -> 8 -> 9) ---
        with tab3:
            st.subheader("💰 Ciclo Final de Pagamento")
            
            # Seção 1: Liquidação (Status 7 -> 8)
            st.markdown("#### ⚖️ 1. Faturas em Liquidação (Retorno do Fiscal)")
            f_status_7 = df[df['status'] == 7].copy()
            if not f_status_7.empty:
                lote_liq = st.multiselect("Selecionar para LIQUIDAR:", f_status_7['nup'].tolist(), key="liq_lote")
                if st.button("✅ Confirmar Liquidação"):
                    for n in lote_liq:
                        mover_status(n, 8)
                        fatura_n = df[df['nup'] == n]['Numero_da_fatura'].values[0]
                        registrar_acao(n, fatura_n, "LIQUIDADO", "Fatura liquidada e pronta para pagamento.")
                    st.rerun()
                st.dataframe(f_status_7[['nup', 'ose', 'valor_liquido']], use_container_width=True)
            else:
                st.info("Nenhuma fatura aguardando liquidação.")

            st.divider()

            # Seção 2: Pagamento (Status 8 -> 9)
            st.markdown("#### 💸 2. Faturas Liquidadas (Prontas para Pagar)")
            f_status_8 = df[df['status'] == 8].copy()
            if not f_status_8.empty:
                lote_pag = st.multiselect("Selecionar para PAGAR (Encerrar):", f_status_8['nup'].tolist(), key="pag_lote")
                if st.button("🏁 Confirmar Pagamento Efetuado"):
                    for n in lote_pag:
                        mover_status(n, 9) # Status 9: Encerrado/Pago
                        fatura_n = df[df['nup'] == n]['Numero_da_fatura'].values[0]
                        registrar_acao(n, fatura_n, "PAGAMENTO_EFETUADO", "Processo encerrado. Pagamento realizado.")
                    st.success("Missão cumprida! Faturas pagas.")
                    time.sleep(1)
                    st.rerun()
                st.dataframe(f_status_8[['nup', 'ose', 'valor_liquido']], use_container_width=True)
            else:
                st.info("Nenhuma fatura pronta para pagamento.")








    elif "FISCALIZAÇÃO" in st.session_state.modulo_ativo or st.session_state.modulo_ativo == "ADMIN":
        st.header("📋 Fiscalização")
        fisc_df = df[df['status'] == 6]
        if st.session_state.user_full_name != "ROSILENE RIBEIRO":
            fisc_df = fisc_df[fisc_df['responsavel_atual'] == st.session_state.user_id]
        st.dataframe(fisc_df[['nup', 'ose', 'ne']])
        n_nf = st.selectbox("NF para", fisc_df['nup'].tolist())
        nf_v = st.text_input("Nº Nota Fiscal")
        if st.button("Enviar p/ Liquidação"):
            cell = aba_p.find(n_nf)
            aba_p.update_cell(cell.row, 15, nf_v)
            mover_status(n_nf, 7)
            st.rerun()

    elif st.session_state.modulo_ativo == "GERENCIAL" or st.session_state.modulo_ativo == "ADMIN":
        st.header("📈 Dashboard")
        st.metric("Economia", f"R$ {pd.to_numeric(df['glosa']).sum():,.2f}")
        st.bar_chart(df['status'].value_counts())

    elif st.session_state.modulo_ativo == "OSE":
        st.header("🏥 Portal OSE")
        st.dataframe(df[df['cnpj'] == st.session_state.user_id][['nup', 'status']])