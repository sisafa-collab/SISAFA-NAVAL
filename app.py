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
caminho_mapeamento = os.path.join(pasta_projeto, "mapeamento-de-processo.png")

st.set_page_config(page_title="SISAFA-NAVAL (HNBra)", layout="centered", page_icon="⚓")

# --- ESTILIZAÇÃO CSS ---
st.markdown("""
    <style>
    [data-testid="stSidebarNav"] {display: none;} 
    .welcome-box { background: #f0f2f6; padding: 15px; border-radius: 10px; border-left: 8px solid #2e6b54; margin-bottom: 25px; font-size: 18px; font-weight: bold; color: #1B3129; }
    .stButton>button { background-color: #2e6b54; color: white; border-radius: 5px; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# --- FUNÇÕES DE CONEXÃO E CACHE (BLINDAGEM DO GOOGLE) ---

@st.cache_resource
def obter_cliente_google():
    """Mantém a sessão com o Google ativa para evitar múltiplos logins"""
    return conectar_google()

@st.cache_data(ttl=60)  # Guarda os dados por 10 minutos
def carregar_dados_cache(nome_aba):
    """Lê dados da planilha e guarda na memória para economizar cota"""
    try:
        # Tenta usar a conexão que já criamos
        client_c = conectar_google()
        if client_c:
            sh_c = client_c.open_by_key(ID_PLANILHA)
            aba = sh_c.worksheet(nome_aba)
            return pd.DataFrame(aba.get_all_records())
    except Exception as e:
        # Se falhar (como a internet do hospital oscilando), avisa mas não trava
        print(f"Erro ao carregar cache da aba {nome_aba}: {e}")
    return pd.DataFrame()

# --- VARIÁVEL GLOBAL PARA ESCRITA (CRUCIAL PARA O LOGIN) ---
try:
    if 'sh' not in locals() or sh is None:
        client_direto = obter_cliente_google()
        if client_direto:
            sh = client_direto.open_by_key(ID_PLANILHA)
            # Aba de processos definida globalmente para mover status
            aba_p = sh.worksheet(ABA_PROCESSOS)
        else:
            sh = None
except Exception as e:
    sh = None
    print(f"Erro ao definir 'sh' global: {e}")

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

def enviar_email_generico(destinatario, assunto, corpo, cc=None):
    """
    Função para enviar e-mails gerais (Solicitação de NF, avisos, etc)
    Ajustada para limpar quebras de linha no Assunto e tratar a lista de CC.
    """
    SMTP_SERVER = "smtp.gmail.com"
    SMTP_PORT = 587
    SMTP_USER = "hnbra.execucaofinanceira@gmail.com"
    SMTP_PASS = st.secrets["smtp_password"] 

    msg = EmailMessage()
    
    # --- CORREÇÃO DO ERRO: Limpeza do Assunto ---
    # O strip() remove espaços e o replace("\n", "") remove o "Enter" que causa o erro
    msg['Subject'] = str(assunto).strip().replace("\n", "").replace("\r", "")
    
    msg['From'] = SMTP_USER
    msg['To'] = destinatario
    
    # --- TRATAMENTO DE CÓPIA (CC) ---
    if cc:
        if isinstance(cc, list):
            # Se for uma lista, remove e-mails vazios e junta com vírgula
            lista_limpa = [e for e in cc if e and str(e).strip()]
            if lista_limpa:
                msg['Cc'] = ", ".join(lista_limpa)
        elif isinstance(cc, str) and cc.strip():
            msg['Cc'] = cc

    # Define o corpo do e-mail
    msg.set_content(corpo)

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.send_message(msg)
        return True
    except Exception as e:
        # Log interno do erro para ajudar no diagnóstico
        print(f"Erro técnico SMTP: {e}")
        st.error(f"Erro no envio técnico: {e}")
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

# --- 2. CONEXÃO GLOBAL E DEFINIÇÃO DE 'sh' ---
try:
    client = conectar_google()
    if client:
        sh = client.open_by_key(ID_PLANILHA)
        aba_p = sh.worksheet(ABA_PROCESSOS)
        df = carregar_dados_cache(ABA_PROCESSOS)
    else:
        sh = None
        df = pd.DataFrame()
except Exception as e:
    st.warning("⚠️ Conexão instável com o Google. Tentando reconectar...")
    sh = None
    df = pd.DataFrame()

# --- CONTROLE DE SESSÃO ---
if 'logged_in' not in st.session_state: 
    st.session_state.logged_in = False
if 'modulo_ativo' not in st.session_state: 
    st.session_state.modulo_ativo = None
# Variáveis de Confirmação (Cura para o AttributeError)
if 'confirmar_secom' not in st.session_state: 
    st.session_state.confirmar_secom = False
if 'confirmar_recebimento' not in st.session_state: 
    st.session_state.confirmar_recebimento = False
if 'confirmar_finalizacao' not in st.session_state: 
    st.session_state.confirmar_finalizacao = False
if 'nups_para_receber' not in st.session_state: st.session_state.nups_para_receber = []


# --- 1. TELA DE LOGIN ---
if not st.session_state.logged_in:
    # O MASCOTE agora está preso aqui dentro. Só aparece se NÃO estiver logado.
    mascote_path = carregar_imagem(caminho_mascote)
    if mascote_path:
        with open(mascote_path, "rb") as f:
            data = base64.b64encode(f.read()).decode()
            st.markdown(f'<img src="data:image/png;base64,{data}" style="position: fixed; bottom: 20px; right: 20px; width: 180px; z-index:999;">', unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        logo_path = carregar_imagem(caminho_logo)
        if logo_path: 
            st.image(logo_path, use_container_width=True)
        else:
            st.markdown("<h1 style='text-align: center;'>⚓ SISAFA-NAVAL</h1>", unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)

        # Formulário de Login
        tipo_acesso = st.radio("Acesso:", ["Interno (NIP)", "Externo (CNPJ)"], horizontal=True)
        u_id = st.text_input(f"Digite seu {'NIP' if 'Interno' in tipo_acesso else 'CNPJ'}")
        senha = st.text_input("Senha", type="password")
        
        if st.button("ACESSAR SISTEMA", use_container_width=True):
            df_users = carregar_dados_cache(ABA_USUARIOS)
            
            if not df_users.empty:
                # --- DINÂMICA DE TAMANHO (NIP 8 | CNPJ 14) ---
                # Identifica o tamanho correto com base na escolha do rádio 'tipo_acesso'
                tamanho_id = 8 if "Interno" in tipo_acesso else 14
                
                # 1. Limpa a coluna da planilha (Primeira Coluna)
                df_users.iloc[:, 0] = (
                    df_users.iloc[:, 0]
                    .astype(str)
                    .str.split('.').str[0]
                    .str.strip()
                    .str.zfill(tamanho_id)
                )
                
                # 2. Limpa o ID digitado pelo usuário
                u_id_limpo = u_id.strip().zfill(tamanho_id)
                
                # Procura o usuário usando os dados vacinados
                user_match = df_users[df_users.iloc[:, 0] == u_id_limpo]
                
                if not user_match.empty:
                    # 1. Pegamos o valor bruto da coluna de senha (Índice 4 = Coluna E)
                    # Forçamos para string e limpamos espaços no início e fim
                    senha_na_planilha = str(user_match.iloc[0, 4]).strip()
                    senha_digitada = senha.strip()

                    # 2. Comparação exata
                    if senha_na_planilha == senha_digitada:
                        st.session_state.logged_in = True
                        st.session_state.user_id = u_id_limpo 
                        st.session_state.user_full_name = str(user_match.iloc[0, 1]).upper()
                        st.session_state.user_perfil = str(user_match.iloc[0, 2]).upper()
                        st.rerun()
                    else:
                        # --- BOX DE DIAGNÓSTICO (Aparecerá apenas se a senha falhar) ---
                        st.error("Senha incorreta.")
                        with st.expander("🔍 Detalhes do erro (Verifique sua Planilha)"):
                            st.write(f"**ID Encontrado:** {u_id_limpo}")
                            st.write(f"**Texto na Planilha:** `{senha_na_planilha}`")
                            st.write(f"**Texto Digitado:** `{senha_digitada}`")
                            st.write(f"**Tamanho Planilha:** {len(senha_na_planilha)} caracteres")
                            st.write(f"**Tamanho Digitado:** {len(senha_digitada)} caracteres")
                            st.info("Dica: Se os tamanhos forem diferentes e o texto parecer igual, há um espaço invisível na sua célula do Google Sheets.")
                else:
                    st.error(f"Usuário {u_id_limpo} não cadastrado.")

# --- 2. TELA DE SELEÇÃO DE MÓDULO (ISSO CURA A TELA BRANCA) ---
elif st.session_state.modulo_ativo is None:
    col_l1, col_l2, col_l3 = st.columns([1.2, 1, 1.2])
    with col_l2:
        if os.path.exists(caminho_logo): st.image(caminho_logo, use_container_width=True)
    
    st.markdown(f"<h1 style='text-align: center; color: #2e6b54;'>⚓ Olá, {st.session_state.user_full_name}</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 20px;'>Selecione o setor de trabalho:</p><br>", unsafe_allow_html=True)
    
    # Cria os botões de SECOM, AUDITORIA, etc., baseado no cadastro do usuário
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
        st.markdown(f"<p style='text-align:center;'><b>ID: {st.session_state.user_id}</b><br>Setor: {st.session_state.modulo_ativo}</p>", unsafe_allow_html=True)
        if st.button("🔄 Trocar de Setor"):
            st.session_state.modulo_ativo = None
            st.rerun()
        if st.button("❌ Sair"):
            st.session_state.logged_in = False
            st.session_state.modulo_ativo = None
            st.rerun()

    st.markdown(f'<div class="welcome-box">⚓ SISAFA-NAVAL: {st.session_state.modulo_ativo}</div>', unsafe_allow_html=True)
    
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
                df_fila['valor_limpo'] = df_fila['valor_apresentado'].apply(limpar_valor)
                df_fila['dt_entrada'] = pd.to_datetime(df_fila['data_entrada'], dayfirst=True, errors='coerce')
                
                hoje = datetime.now()
                df_fila['dias_fila'] = (hoje - df_fila['dt_entrada']).dt.days

                # Cálculos de Temporalidade
                aceitavel = len(df_fila[df_fila['dias_fila'] <= 1])
                atencao = len(df_fila[(df_fila['dias_fila'] >= 2) & (df_fila['dias_fila'] <= 4)])
                atraso = len(df_fila[df_fila['dias_fila'] > 5])
                
                # --- INTERFACE DE INDICADORES (KPIs) ---
                st.markdown("### 📊 Situação da Fila (Faturas SECOM)")
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Total de Faturas", f"{len(df_fila)}")
                c2.metric("🟢 Aceitável (até 1d)", f"{aceitavel}")
                c3.metric("🟡 Atenção (2-4d)", f"{atencao}")
                c4.metric("🔴 Em Atraso (>5d)", f"{atraso}")

                with st.expander("💰 Detalhamento por Competência", expanded=False):
                    resumo_comp = df_fila.groupby(['mes_sigla', 'ano_competencia'])['valor_limpo'].sum().reset_index()
                    resumo_comp.columns = ['Mês', 'Ano', 'Total (R$)']
                    st.table(resumo_comp.style.format({'Total (R$)': 'R$ {:,.2f}'}))
                
                st.divider()

            # --- TABELA DE PROCESSOS E RECEBIMENTO ---
            st.subheader("📥 Processos aguardando Auditoria")
            
            if df_fila.empty:
                st.info("Não há faturas na fila no momento.")
            else:
                # Exibição da Tabela
                st.dataframe(
                    df_fila[['nup', 'ose', 'valor_apresentado', 'mes_sigla', 'ano_competencia', 'dias_fila']], 
                    use_container_width=True
                )
                
                # Seleção e Ação de Recebimento Direto
                nups_sel = st.multiselect("Selecione os NUPs para trazer para sua mesa:", df_fila['nup'].tolist(), key="sel_fila_direto")
                
                if st.button("📥 RECEBER PROCESSOS SELECIONADOS", use_container_width=True):
                    if nups_sel:
                        with st.spinner("Movimentando para auditagem..."):
                            for n in nups_sel:
                                mover_status(n, 2, auditor_nip=st.session_state.user_id)
                                try:
                                    fat_n = df[df['nup'] == n]['Numero_da_fatura'].values[0]
                                    registrar_acao(n, fat_n, "RECEBIMENTO", f"Auditor {st.session_state.user_id} recebeu.")
                                except: pass
                        st.toast("Sucesso! Processos movidos para 'Em Auditagem'.", icon="✅")
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.warning("⚠️ Selecione ao menos um NUP.")

                st.divider()

                # --- FERRAMENTA DE CORREÇÃO (ERRO HUMANO) ---
                with st.expander("🛠️ CORRIGIR ERROS DE CADASTRO (NUP, Fatura ou Valor)"):
                    st.write("Selecione um processo da fila para editar os dados originais:")
                    
                    nup_edit = st.selectbox("Escolha o NUP para corrigir:", [""] + df_fila['nup'].tolist(), key="sb_edit_fila")
                    
                    if nup_edit:
                        dados = df_fila[df_fila['nup'] == nup_edit].iloc[0]
                        
                        col_e1, col_e2, col_e3 = st.columns(3)
                        novo_nup = col_e1.text_input("Novo NUP:", value=dados['nup'])
                        nova_fat = col_e2.text_input("Nova Fatura:", value=dados['Numero_da_fatura'])
                        novo_val = col_e3.number_input("Novo Valor (R$):", value=float(dados['valor_limpo']), format="%.2f")
                        
                        if st.button("💾 SALVAR CORREÇÃO", use_container_width=True):
                            try:
                                # Acessa a planilha para editar a célula exata
                                aba_edit = sh.worksheet(ABA_PROCESSOS)
                                celula = aba_edit.find(nup_edit)
                                
                                if celula:
                                    # Colunas: B=2 (NUP), E=5 (Fatura), F=6 (Valor), H=8 (V. Líquido)
                                    aba_edit.update_cell(celula.row, 2, novo_nup)
                                    aba_edit.update_cell(celula.row, 5, nova_fat)
                                    aba_edit.update_cell(celula.row, 6, novo_val)
                                    aba_edit.update_cell(celula.row, 8, novo_val)
                                    
                                    registrar_acao(novo_nup, nova_fat, "CORREÇÃO", f"Corrigido por {st.session_state.user_id}")
                                    
                                    st.success("✅ Dados atualizados!")
                                    st.cache_data.clear() # Limpa o cache para atualizar a tabela
                                    time.sleep(1)
                                    st.rerun()
                            except Exception as e:
                                st.error(f"Erro ao salvar: {e}")

        # 2. ABA: EM AUDITAGEM
        with t_mesa:
            # --- CÁLCULO DOS INDICADORES TÉCNICOS (Status 2) ---
            df_total_auditagem = df[df['status'] == 2].copy()
            
            if not df_total_auditagem.empty:
                # 1. Preparação dos dados
                df_total_auditagem['valor_limpo'] = df_total_auditagem['valor_apresentado'].apply(limpar_valor)
                df_total_auditagem['dt_mov'] = pd.to_datetime(df_total_auditagem.iloc[:, 13], dayfirst=True, errors='coerce')
                
                hoje = datetime.now()
                df_total_auditagem['dias_auditoria'] = (hoje - df_total_auditagem['dt_mov']).dt.days
                
                # --- INTERFACE DE INDICADORES (KPIs) ---
                st.markdown("### 📊 Situação Geral das Faturas em Auditagem")
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Total em Auditagem", f"{len(df_total_auditagem)}")
                c2.metric("🟢 Aceitável", f"{len(df_total_auditagem[df_total_auditagem['dias_auditoria'] <= 10])}")
                c3.metric("🟡 Atenção", f"{len(df_total_auditagem[(df_total_auditagem['dias_auditoria'] > 10) & (df_total_auditagem['dias_auditoria'] <= 15)])}")
                c4.metric("🔴 Em Atraso", f"{len(df_total_auditagem[df_total_auditagem['dias_auditoria'] > 15])}")

                st.divider()

            # --- MESA DE TRABALHO (VISÃO COLETIVA) ---
            st.subheader("🩺 Mesa de Trabalho da Auditoria")
            df_mesa = df[df['status'] == 2].copy()

            if df_mesa.empty:
                st.info("Não há processos em auditagem no momento.")
            else:
                st.write("**Faturas em análise técnica no setor:**")
                st.dataframe(df_mesa[['nup', 'ose', 'valor_apresentado', 'mes_sigla', 'ano_competencia', 'obs']], use_container_width=True)
                
                st.divider()
                
                nup_audit = st.selectbox("Selecione o NUP para realizar a análise:", [""] + df_mesa['nup'].tolist(), key="sb_nup_analise_mesa_final")
                
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

                    # --- CONFERÊNCIA E ENVIO (SEM REDUNDÂNCIA) ---
                    st.markdown("### 🛡️ Conferência e Finalização")
                    with st.container(border=True):
                        try:
                            cnpj_fatura = str(dados_nup['cnpj']).strip().split('.')[0]
                            df_ose_info = carregar_dados_cache(ABA_TABELA_A)
                            linha_ose = df_ose_info[df_ose_info['CNPJ'].astype(str).str.contains(cnpj_fatura)]
                            df_users = carregar_dados_cache(ABA_USUARIOS)
                            match_user = df_users[df_users['NIP'].astype(str).str.strip() == str(st.session_state.user_id).strip()]
                            
                            if not match_user.empty and not linha_ose.empty:
                                email_auditor = match_user['Email'].values[0] if 'Email' in match_user.columns else match_user['E-mail'].values[0]
                                email_destino = linha_ose['E-mail Principal da OSE'].values[0]
                                nome_ose_oficial = linha_ose['Razão Social'].values[0]
                                
                                st.write(f"🏢 **OSE:** {nome_ose_oficial} | 📩 **E-mail:** {email_destino}")
                                trava_confirmacao = st.checkbox("Confirmo que os dados estão corretos para envio/finalização.")
                            else:
                                trava_confirmacao = False
                                st.error("⚠️ Dados de e-mail não localizados na base.")
                        except:
                            trava_confirmacao = False

                    col_fin, col_mail = st.columns(2)
                    
                    # AÇÃO DIRETA: FINALIZAR
                    if col_fin.button("✅ FINALIZAR AUDITORIA", use_container_width=True):
                        if glosa_input > 0 and not just_glosa:
                            st.error("⚠️ Justificativa obrigatória para glosa.")
                        elif not trava_confirmacao:
                            st.warning("⚠️ Marque a caixa de conferência acima.")
                        else:
                            with st.spinner("Finalizando análise..."):
                                mover_status(nup_audit, 3, valor_glosa=glosa_input, valor_liq=v_liquido, obs_texto=just_glosa)
                                registrar_acao(nup_audit, num_fat, "AUDITORIA_CONCLUIDA", f"V.Líq: {v_liquido}")
                                st.success("✅ Auditoria finalizada e movida para o Fiscal!")
                                time.sleep(1.5)
                                st.rerun()

                    # AÇÃO DIRETA: ENVIAR E-MAIL
                    if col_mail.button("📧 ENCAMINHAR GLOSA P/ OSE", use_container_width=True, disabled=not trava_confirmacao):
                        with st.spinner("Enviando comunicado..."):
                            if disparar_email_glosa(email_destino, num_fat, glosa_input, just_glosa, nome_ose_oficial, email_auditor):
                                registrar_acao(nup_audit, num_fat, "EMAIL_GLOSA_ENVIADO", f"Destino: {email_destino}")
                                st.toast("E-mail enviado!", icon="✅")
                            else:
                                st.error("Falha no envio do e-mail.")


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
            "📥 Caixa de entrada", "📄 Gestão de NE", "💸 Gestão de Pagamentos", 
            "📊 Estatísticas e Indicadores", "🔍 Consultas", "🤝 Relacionamento"
        ])

        # --- ABA 1: CAIXA DE ENTRADA ---
        with tab1:
            st.header("📥 Caixa de Entrada")
            
            meses_siglas = {
                1:'JAN', 2:'FEV', 3:'MAR', 4:'ABR', 5:'MAI', 6:'JUN',
                7:'JUL', 8:'AGO', 9:'SET', 10:'OUT', 11:'NOV', 12:'DEZ'
            }

            # =================================================================
            # 1. PARTE: FATURAS AUDITADAS (Vindo da Auditoria - Status 3)
            # =================================================================
            st.subheader("📥 Faturas auditadas aguardando Recebimento no Financeiro")
            
            df_fila_aud = df[df['status'] == 3].copy()
            
            # Cruzamento com Logs para pegar data de envio da Auditoria
            try:
                logs_raw = aba_l.get_all_records()
                df_logs = pd.DataFrame(logs_raw)
                df_envio_aud = df_logs[df_logs['acao'] == "ENCAMINHADO_PARA_FINANCEIRO"].copy()
                df_envio_aud['dt_chegada'] = pd.to_datetime(df_envio_aud['data_hora'], dayfirst=True, errors='coerce')
                df_envio_aud = df_envio_aud.sort_values('dt_chegada', ascending=False).drop_duplicates('nup')
                df_fila_aud = df_fila_aud.merge(df_envio_aud[['nup', 'dt_chegada']], on='nup', how='left')
            except:
                df_fila_aud['dt_chegada'] = datetime.now()

            if df_fila_aud.empty:
                st.info("Nenhuma fatura vinda da Auditoria no momento.")
            else:
                df_fila_aud['mes_sigla'] = pd.to_numeric(df_fila_aud['mes_competencia'], errors='coerce').map(meses_siglas)
                hoje = datetime.now()
                df_fila_aud['dias_espera'] = (hoje - df_fila_aud['dt_chegada']).dt.days.fillna(0).astype(int)

                st.dataframe(
                    df_fila_aud[['nup','ose','valor_liquido','mes_sigla','ano_competencia','dias_espera']],
                    use_container_width=True,
                    key="df_aud_recepcao"
                )
                
                nups_aud_sel = st.multiselect("Selecionar faturas auditadas para receber:", df_fila_aud['nup'].tolist(), key="ms_aud_recep")
                
                if st.button("✅ Receber Faturas Auditadas", key="btn_aud_recep"):
                    if nups_aud_sel:
                        with st.spinner("Recebendo..."):
                            for n in nups_aud_sel:
                                mover_status(n, 4) # Evolui para Aguard. NE
                                fat_n = df[df['nup'] == n]['Numero_da_fatura'].values[0]
                                registrar_acao(n, fat_n, "RECEBIMENTO_FINANCEIRO", "Fatura auditada recebida pela Execução.")
                        st.success(f"{len(nups_aud_sel)} faturas recebidas!")
                        time.sleep(1)
                        st.rerun()

            st.divider()

            # =================================================================
            # 2. PARTE: NOTAS FISCAIS CERTIFICADAS (Vindo do Fiscal - Status 6)
            # =================================================================
            st.subheader("📥 Notas Fiscais certificadas aguardando Recebimento no Financeiro")
            
            # Filtramos processos que o Fiscal já digitou a NF (Status 6)
            df_fila_fiscal = df[
            (df['status'] == 6) & 
            (df['nf'].astype(str).str.strip() != "") & 
            (df['nf'].notna())
            ].copy()
            
            if df_fila_fiscal.empty:
                st.info("Nenhuma Nota Fiscal aguardando conferência do Fiscal.")
            else:
                df_fila_fiscal['mes_sigla'] = pd.to_numeric(df_fila_fiscal['mes_competencia'], errors='coerce').map(meses_siglas)
                
                # Exibição da tabela com a coluna 'nf' (Coluna onde o fiscal digita)
                # Assumindo que a coluna na sua planilha chama-se 'nf'
                cols_nf = ['nup','cnpj','ose','mes_sigla','ano_competencia','valor_liquido', 'ne','nf']
                st.dataframe(df_fila_fiscal[cols_nf].rename(columns={'nf': 'Número da NF'}), use_container_width=True)
                
                # Seleção por NOTA FISCAL
                lista_nfs = sorted(df_fila_fiscal['nf'].unique().tolist())
                nfs_sel = st.multiselect("Selecione a(s) Nota(s) Fiscal(is) para aceitar:", options=lista_nfs, key="ms_nf_recep")
                
                if st.button("🚀 Aceitar e Liquidar Notas Fiscais", key="btn_nf_recep"):
                    if nfs_sel:
                        # Buscamos todos os NUPs vinculados a essas NFs
                        nups_da_nf = df_fila_fiscal[df_fila_fiscal['nf'].isin(nfs_sel)]['nup'].tolist()
                        
                        with st.spinner(f"Processando {len(nups_da_nf)} faturas..."):
                            for n in nups_da_nf:
                                mover_status(n, 7) # Evolui de 6 para 7 (Liquidando/Retorno do Fiscal)
                                fat_n = df[df['nup'] == n]['Numero_da_fatura'].values[0]
                                nf_n = df_fila_fiscal[df_fila_fiscal['nup'] == n]['nf'].values[0]
                                registrar_acao(n, fat_n, "NF_ACEITA_FINANCEIRO", f"NF {nf_n} conferida e aceita pela Execução.")
                        
                        st.success(f"✅ {len(nfs_sel)} Notas Fiscais aceitas! Processos movidos para liquidação.")
                        time.sleep(1.5)
                        st.rerun()
                    else:
                        st.warning("⚠️ Selecione ao menos uma NF para aceitar.")

        # --- ABA 2: GESTÃO DE NE (Status 4 -> 5 -> 6) ---
        with tab2:
            # --- SEÇÃO 1: EMISSÃO DE NE ---
            st.markdown("### 📝 1. Emitir Nota de Empenho (NE)")
            
            meses_siglas = {
                1:'JAN', 2:'FEV', 3:'MAR', 4:'ABR', 5:'MAI', 6:'JUN',
                7:'JUL', 8:'AGO', 9:'SET', 10:'OUT', 11:'NOV', 12:'DEZ'
            }
            
            f_status_4 = df[df['status'] == 4].copy()
            
            if not f_status_4.empty:
                f_status_4['mes_sigla'] = pd.to_numeric(f_status_4['mes_competencia'], errors='coerce').map(meses_siglas)

                nups_sel = st.multiselect(
                    "Selecione o(s) NUP(s) para empenhar (Devem ser da mesma empresa):", 
                    f_status_4['nup'].tolist(), 
                    key="sel_ne_batch"
                )
                
                trava_cnpj = False
                if nups_sel:
                    df_conf = f_status_4[f_status_4['nup'].isin(nups_sel)].copy()
                    df_conf['v_liq_num'] = df_conf['valor_liquido'].apply(limpar_valor)
                    lista_cnpjs = df_conf['cnpj'].unique()
                    trava_cnpj = len(lista_cnpjs) > 1
                    
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

                col_input1, col_input2 = st.columns([1,1])
                with col_input1:
                    cod_ne_final = st.text_input("Número Final da NE (ex: 00052)", key="input_ne_num")
                
                if st.button("🚀 Cadastrar NE", disabled=trava_cnpj or not nups_sel, use_container_width=True):
                    if not cod_ne_final:
                        st.warning("⚠️ Digite o número da NE antes de prosseguir.")
                    else:
                        ne_completa = f"78770000001{datetime.now().year}NE{cod_ne_final}"
                        cnpj_alvo = f_status_4[f_status_4['nup'].isin(nups_sel)]['cnpj'].iloc[0]
                        
                        with st.spinner(f"Gravando NE {ne_completa}..."):
                            for nup in nups_sel:
                                cell = aba_p.find(nup)
                                if cell:
                                    aba_p.update_cell(cell.row, 15, ne_completa)
                                    mover_status(nup, 5)
                                    fatura_n = df[df['nup'] == nup]['Numero_da_fatura'].values[0]
                                    registrar_acao(nup, fatura_n, "NE_CADASTRADA", f"NE {ne_completa} vinculada ao CNPJ {cnpj_alvo}")
                            
                            st.success(f"✅ Sucesso! NE {ne_completa} cadastrada.")
                            time.sleep(1.5)
                            st.rerun()

                st.divider()
                st.subheader("📋 Processos Disponíveis para Empenho")
                cols_v = ['nup','cnpj','ose','mes_sigla','ano_competencia','valor_liquido']
                st.dataframe(f_status_4[cols_v].rename(columns={'mes_sigla':'Mês'}), use_container_width=True)
            else:
                st.info("Não há faturas aguardando emissão de NE.")

            st.markdown("---") # Separador visual dentro da aba

            # --- SEÇÃO 2: ENVIO PARA FISCALIZAÇÃO ---
            st.markdown("### 📤 2. Encaminhar para Fiscalização (Por Nota de Empenho)")
        
            f_status_5 = df[df['status'] == 5].copy()
        
            if not f_status_5.empty:
            # 1. BUSCA DE FISCAIS (Fiscalização de contrato e FISCAL_GLOBAL)
                df_users_fiscal = carregar_dados_cache(ABA_USUARIOS)
                if not df_users_fiscal.empty:
                # Filtra perfis (considerando que perfil está na coluna de índice 2)
                # O lower() e strip() servem para evitar erros de digitação na planilha
                    fiscais_disp = df_users_fiscal[
                        df_users_fiscal.iloc[:, 2].str.lower().str.contains("fiscalização de contrato|fiscal_global", na=False)
                    ]
                # Criamos a lista de nomes (Coluna de índice 1)
                    lista_fiscais = sorted(fiscais_disp.iloc[:, 1].unique().tolist())
                else:
                    lista_fiscais = []

            # Interface de seleção
                lista_nes_disponiveis = sorted(f_status_5['ne'].unique().tolist())

                c_ne, c_fisc = st.columns([2, 1])
                with c_ne:
                    selecao_ne = st.multiselect(
                        "Selecione a(s) Nota(s) de Empenho:",
                        options=lista_nes_disponiveis,
                        key="multisel_envio_fiscal_ne"
                    )
                with c_fisc:
                    fiscal_destinatario = st.selectbox(
                        "Enviar para qual Fiscal?",
                        options=[""] + lista_fiscais,
                        help="O nome selecionado será gravado no histórico do processo.",
                        key="sb_fiscal_destino"
                    )

                if st.button("📧 Encaminhar Empenhos p/ Fiscalização", use_container_width=True):
                    if selecao_ne and fiscal_destinatario:
                        nups_para_enviar = f_status_5[f_status_5['ne'].isin(selecao_ne)]['nup'].tolist()
                    
                        with st.spinner(f"Encaminhando {len(nups_para_enviar)} faturas..."):
                            for nup in nups_para_enviar:
                                # 1. Evolui para Status 6
                                mover_status(nup, 6)
                            
                            # 2. Busca dados para o Log e Histórico
                                dados_n = f_status_5[f_status_5['nup'] == nup].iloc[0]
                                fat_n = dados_n['Numero_da_fatura']
                                v_apres = dados_n['valor_apresentado']
                            
                                # 3. Log da ação (Micro)
                                registrar_acao(nup, fat_n, "ENVIO_FISCALIZACAO", f"Enviado p/ Fiscal: {fiscal_destinatario}")
                            
                            # 4. Histórico (Macro) - Agora com o nome do Fiscal no campo OBS
                            # mover_status(nup, 6) no SISAFA costuma registrar o histórico, 
                            # mas chamamos o registrar_historico aqui para forçar a OBS personalizada:
                                registrar_historico(
                                    nup, 
                                    fat_n, 
                                    "5", 
                                    "6", 
                                    v_apres, 
                                    f"Empenho {dados_n['ne']} enviado para: {fiscal_destinatario}"
                                )
                    
                        st.success(f"✅ Sucesso! Empenhos enviados para {fiscal_destinatario}.")
                        time.sleep(1.5)
                        st.rerun()
                    elif not fiscal_destinatario:
                        st.warning("⚠️ Por favor, selecione o Fiscal de destino.")
                    else:
                        st.warning("⚠️ Selecione ao menos uma Nota de Empenho.")

            # Tabela de visualização (abaixo das seleções)
                st.subheader("📊 Faturas Empenhadas aguardando envio")
                f_status_5['mes_sigla'] = f_status_5['mes_competencia'].map(meses_siglas)
                cols_f = ['ne', 'ose', 'nup', 'valor_liquido', 'mes_sigla', 'ano_competencia']
                st.dataframe(f_status_5[cols_f].sort_values(by='ne'), use_container_width=True)
            else:
                st.info("Não há Notas de Empenho aguardando envio.")

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
                st.dataframe(f_status_7[['nup', 'ose', 'valor_liquido','nf']], use_container_width=True)
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
                st.dataframe(f_status_8[['nup', 'ose', 'valor_liquido','nf']], use_container_width=True)
            else:
                st.info("Nenhuma fatura pronta para pagamento.")

        # --- ABA 4: ESTATÍSTICAS E INDICADORES (KPIs Financeiros) ---
        with tab4:
            st.header("📊 Estatística e Indicadores")
            
            # Filtros Rápidos
            col_f1, col_f2 = st.columns(2)
            anos_disp = sorted(df['ano_competencia'].unique(), reverse=True)
            ano_sel = col_f1.selectbox("Filtrar por Ano:", ["Todos"] + list(anos_disp), key="f_ano_exec")
            
            df_e = df.copy()
            if ano_sel != "Todos": df_e = df_e[df_e['ano_competencia'] == ano_sel]
            
            df_e['v_ap_num'] = df_e['valor_apresentado'].apply(limpar_valor)
            df_e['v_liq_num'] = df_e['valor_liquido'].apply(limpar_valor)
            df_e['glosa_num'] = df_e['glosa'].apply(limpar_valor)

            # --- 1. MÉTRICAS GLOBAIS ---
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Processos no Ciclo", len(df_e))
            c2.metric("Apresentado", f"R$ {df_e['v_ap_num'].sum():,.2f}")
            c3.metric("Economia (Glosa)", f"R$ {df_e['glosa_num'].sum():,.2f}")
            c4.metric("Líquido a Pagar", f"R$ {df_e['v_liq_num'].sum():,.2f}")

            st.divider()
            
            # --- 2. GRÁFICOS DE DESEMPENHO ---
            cg1, cg2 = st.columns(2)
            
            with cg1:
                # Distribuição por Fase do Ciclo
                status_map = {
                    1: "Fila Auditoria", 2: "Em Auditoria", 3: "Fila Execução",
                    4: "Aguard. NE", 5: "Empenhado", 6: "Fiscalização",
                    7: "Liquidando", 8: "Pronto Pagar", 9: "Pago/Encerrado"
                }
                status_counts = df_e['status'].map(status_map).value_counts().reset_index()
                st.plotly_chart(px.pie(status_counts, values='count', names='status', title="Distribuição por Fase", hole=0.4), use_container_width=True)
            
            with cg2:
                # Top 10 OSEs por Valor Líquido
                top_ose_exec = df_e.groupby('ose')['v_liq_num'].sum().sort_values(ascending=False).head(10).reset_index()
                st.plotly_chart(px.bar(top_ose_exec, x='v_liq_num', y='ose', orientation='h', title="Top 10 OSEs (Valor Líquido)", color_discrete_sequence=['#2e6b54']), use_container_width=True)

        # --- ABA 5: CONSULTAS (Rastreabilidade Total) ---
        with tab5:
            st.subheader("🔍 Consultas")
            
            termo_busca_exec = st.text_input(
                "Pesquise por NUP, OSE, CNPJ, Nota de Empenho ou Fatura:", 
                placeholder="Ex: 00052, HNBra, 63060...",
                key="busca_global_exec"
            )

            if termo_busca_exec:
                # Busca em múltiplas colunas simultaneamente
                mask_exec = (
                    df['nup'].astype(str).str.contains(termo_busca_exec, case=False, na=False) |
                    df['ose'].astype(str).str.contains(termo_busca_exec, case=False, na=False) |
                    df['ne'].astype(str).str.contains(termo_busca_exec, case=False, na=False) |
                    df['Numero_da_fatura'].astype(str).str.contains(termo_busca_exec, case=False, na=False)
                )
                res_exec = df[mask_exec]

                if res_exec.empty:
                    st.warning("Nenhum registro encontrado.")
                else:
                    st.write(f"📂 **{len(res_exec)}** resultados encontrados:")
                    st.dataframe(res_exec[['nup', 'ose', 'Numero_da_fatura', 'ne', 'status']], use_container_width=True)
                    
                    nup_detalhe = st.selectbox("Selecione o NUP para ver o histórico completo:", [""] + res_exec['nup'].tolist(), key="sb_detalhe_exec")

                    if nup_detalhe:
                        # Snapshot
                        dados_nup = df[df['nup'] == nup_detalhe].iloc[0]
                        st.info(f"📍 **Status Atual:** Fase {dados_nup['status']} | **Responsável:** {dados_nup['responsavel_atual']} | **NE:** {dados_nup['ne']}")

                        # Histórico de Movimentações (Status)
                        st.markdown("### 👣 Movimentações de Status")
                        try:
                            aba_h = sh.worksheet(ABA_HISTORICO)
                            df_h = pd.DataFrame(aba_h.get_all_records())
                            track = df_h[df_h['nup'] == nup_detalhe].sort_values(by='timestamp')
                            if not track.empty:
                                for _, r in track.iterrows():
                                    st.caption(f"🕒 {r['timestamp']} | **{r['status_origem']}** ⮕ **{r['status_destino']}** (Usuário: {r['usuario']})")
                            else: st.write("Sem histórico de movimentação.")
                        except: st.error("Erro ao carregar histórico.")

                        # Logs de Ações (E-mails, Recebimentos, NEs)
                        st.markdown("### 📝 Logs de Eventos")
                        try:
                            aba_l = sh.worksheet(ABA_LOGS_ACOES)
                            df_l = pd.DataFrame(aba_l.get_all_records())
                            logs = df_l[df_l['nup'] == nup_detalhe].sort_values(by='data_hora', ascending=False)
                            if not logs.empty:
                                st.table(logs[['data_hora', 'acao', 'militar_nip', 'detalhes']])
                            else: st.write("Nenhuma ação específica registrada.")
                        except: st.error("Erro ao carregar logs.")

        # --- ABA 6: RELACIONAMENTO (Inbox OSE) ---
        with tab6:
            st.subheader("🤝 Relacionamento")
            st.write("Dúvidas financeiras e questionamentos de faturas enviados pelas OSEs.")

            try:
                aba_msg = sh.worksheet(ABA_MENSAGENS)
                df_msg = pd.DataFrame(aba_msg.get_all_records())
                
                if df_msg.empty:
                    st.info("Nenhuma mensagem pendente.")
                else:
                    pendentes = len(df_msg[df_msg['status_resposta'] == 'PENDENTE'])
                    st.metric("Mensagens Pendentes", pendentes)
                    
                    df_exibir = df_msg.sort_values(by='status_resposta', ascending=False)
                    st.dataframe(df_exibir[['nup', 'cnpj_ose', 'assunto', 'status_resposta']], use_container_width=True)

                    st.markdown("---")
                    nup_msg = st.selectbox("Selecione o NUP para responder à OSE:", [""] + df_msg['nup'].tolist(), key="sel_msg_exec")

                    if nup_msg:
                        item = df_msg[df_msg['nup'] == nup_msg].iloc[0]
                        with st.chat_message("user"):
                            st.write(f"**Assunto:** {item['assunto']}")
                            st.write(item['mensagem_corpo'])
                        
                        resp_exec = st.text_area("Parecer da Execução Financeira:", placeholder="Digite aqui a resposta para a OSE...")
                        if st.button("📤 Enviar Parecer Financeiro"):
                            if resp_exec:
                                registrar_acao(nup_msg, "N/A", "RESPOSTA_FINANCEIRA", "Financeiro respondeu questionamento da OSE.")
                                st.success("Resposta enviada para o portal da OSE!")
                                time.sleep(1)
                                st.rerun()
            except Exception as e:
                st.error(f"Erro na aba relacionamento: {e}")

    # =================================================================
    # MÓDULO 4: FISCALIZAÇÃO DE CONTRATOS (Ajustado para FISCAL/FISCAL_GLOBAL)
    # =================================================================
    elif "FISCAL" in st.session_state.modulo_ativo or st.session_state.modulo_ativo == "ADMIN":
        st.header("📋 Fiscalização de Contratos em Saúde")

        # Definição das 3 abas solicitadas
        tab_visao, tab_nf, tab_rel = st.tabs([
            "🔭 Visão Geral", "🧾 Empenhos aguardando NF", "💬 Relacionamento"
        ])

        # --- LÓGICA DE PERMISSÃO ---
        # Garantimos que o NIP logado tenha sempre 8 dígitos (preenchendo com zero se necessário)
        user_nip = str(st.session_state.user_id).strip().zfill(8)
        is_global = (user_nip == "95039023") # Rosilene

        mapa_status_fisc = {
            1: "1 - FATURA CADASTRADA", 2: "2 - EM AUDITAGEM", 3: "3 - AUDITADA",
            4: "4 - AGUARDANDO EMISSÃO DE NE", 5: "5 - FATURA EMPENHADA",
            6: "6 - AGUARDANDO EMISSÃO DE NF", 7: "7 - EM LIQUIDAÇÃO",
            8: "8 - FATURA LIQUIDADA", 9: "9 - FATURA PAGA"
        }

        # 1. ABA: VISÃO GERAL
        with tab_visao:
            # --- 1. INCLUSÃO DA IMAGEM FIXA ---
            mapeamento_path = carregar_imagem(caminho_mapeamento)
            
            if mapeamento_path:
                with open(mapeamento_path, "rb") as f:
                    data = base64.b64encode(f.read()).decode()
                    st.markdown(
                        f'<img src="data:image/png;base64,{data}" '
                        'style="position: fixed; bottom: 20px; right: 20px; width: 220px; z-index:998; opacity: 0.9; pointer-events: none;">',
                        unsafe_allow_html=True
                    )
            else:
                # Se a imagem não aparecer, isso vai te dizer o porquê (apenas para você ver)
                st.sidebar.error(f"Arquivo não encontrado: {caminho_mapeamento}")
                
            st.subheader("Meus contratos")
            df_ose_master = carregar_dados_cache(ABA_TABELA_A)
            
            if df_ose_master.empty:
                st.error("Não foi possível carregar a Tabela-A. Verifique a conexão.")
            else:
                # Padronização das colunas
                df_ose_master.columns = [c.strip().replace(' ', '_').upper() for c in df_ose_master.columns]
                
                col_nip = "NIP_DO_GESTOR_TITULAR" 
                col_nip_sub = "NIP_DO_GESTOR_SUBSTITUTO"

                if is_global:
                    st.success(f"🔓 Perfil Global: Rosilene ({user_nip})")
                    df_fiscal = df_ose_master.copy()
                else:
                    # --- A "VACINA" DO ZERO À ESQUERDA ---
                    # Verificamos se a coluna principal existe
                    if col_nip in df_ose_master.columns:
                        # Limpamos a coluna do Titular
                        df_ose_master[col_nip] = df_ose_master[col_nip].apply(
                            lambda x: str(x).split('.')[0].strip().zfill(8)
                        )
                        
                        # Criamos o filtro base (Titular)
                        filtro = (df_ose_master[col_nip] == user_nip)
                        
                        # Se existir a coluna de Substituto, limpamos e adicionamos ao filtro
                        if col_nip_sub in df_ose_master.columns:
                            df_ose_master[col_nip_sub] = df_ose_master[col_nip_sub].apply(
                                lambda x: str(x).split('.')[0].strip().zfill(8)
                            )
                            # Filtro: Titular OU Substituto
                            filtro = filtro | (df_ose_master[col_nip_sub] == user_nip)
                        
                        df_fiscal = df_ose_master[filtro].copy()
                    else:
                        st.error(f"❌ Coluna '{col_nip}' não encontrada na Tabela-A.")
                        df_fiscal = pd.DataFrame()

                # --- EXIBIÇÃO ---
                if df_fiscal.empty:
                    st.warning(f"⚠️ Nenhum contrato vinculado ao NIP {user_nip}.")
                else:
                    st.write("**Contratos sob sua responsabilidade (Titular ou Substituto):**")
                    st.dataframe(df_fiscal[['CNPJ', 'RAZÃO_SOCIAL']], use_container_width=True)
                    st.divider()
                    
                    st.subheader("Situação geral")
                    ose_sel = st.selectbox("Selecione a Organização:", [""] + df_fiscal['RAZÃO_SOCIAL'].tolist(), key="fisc_sel_v2")
                    
                    if ose_sel:
                        # Pegamos o CNPJ da empresa escolhida
                        cnpj_ose = str(df_fiscal[df_fiscal['RAZÃO_SOCIAL'] == ose_sel]['CNPJ'].iloc[0]).split('.')[0].strip()
                        
                        # Filtramos na base principal de processos
                        df_proc_fisc = df[df['cnpj'].astype(str).str.contains(cnpj_ose)].copy()
                        
                        if not df_proc_fisc.empty:
                            df_proc_fisc['situação_texto'] = df_proc_fisc['status'].map(mapa_status_fisc)
                            st.write(f"📋 **Processos de {ose_sel}:**")
                            st.dataframe(df_proc_fisc[['nup', 'Numero_da_fatura', 'situação_texto']].rename(columns={'situação_texto': 'Situação'}), use_container_width=True)
                            
                            # --- DASHBOARD ---
                            c1, c2 = st.columns([1, 1])
                            with c1:
                                df_pizza = df_proc_fisc['situação_texto'].value_counts().reset_index()
                                fig = px.pie(df_pizza, values='count', names='situação_texto', title="Status dos Processos", hole=0.4)
                                st.plotly_chart(fig, use_container_width=True)
                            with c2:
                                df_proc_fisc['v_liq'] = df_proc_fisc['valor_liquido'].apply(limpar_valor)
                                tramito_val = df_proc_fisc[df_proc_fisc['status'] < 9]['v_liq'].sum()
                                st.metric("Total em Trâmite", f"R$ {tramito_val:,.2f}")

        # 2. ABA: EMPENHOS AGUARDANDO NF
        with tab_nf:
            st.markdown("### 🧾 NOTAS DE EMPENHO AGUARDANDO EMISSÃO DE NOTA FISCAL")
            
            # --- 1. DEFINIÇÃO DA BASE TOTAL (Evita o NameError) ---
            df_s6_total = df[df['status'] == 6].copy()
            
            # --- 2. LÓGICA DE PRIVACIDADE ---
            if is_global:
                # Rosilene vê tudo
                df_s6 = df_s6_total
            else:
                # Fiscais comuns só vêem o que está no df_fiscal (criado na Aba 1)
                if 'df_fiscal' in locals() and not df_fiscal.empty:
                    # Extraímos os CNPJs do fiscal (limpando para bater com a base)
                    meus_cnpjs = df_fiscal['CNPJ'].astype(str).str.split('.').str[0].unique().tolist()
                    # Filtramos: CNPJ do processo deve estar na lista do fiscal
                    df_s6 = df_s6_total[df_s6_total['cnpj'].astype(str).str.contains('|'.join(meus_cnpjs), na=False)].copy()
                else:
                    # Se não tem contratos vinculados, não vê nada
                    df_s6 = pd.DataFrame()

            # --- 3. EXIBIÇÃO E INTERFACE ---
            if df_s6.empty:
                st.info("Não há Notas de Empenho aguardando NF para os seus contratos.")
            else:
                # Indicadores de Prazo (Coluna 14 / Índice 13)
                df_s6['dt_mov'] = pd.to_datetime(df_s6.iloc[:, 13], dayfirst=True, errors='coerce')
                df_s6['dias'] = (datetime.now() - df_s6['dt_mov']).dt.days.fillna(0).astype(int)
                
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Minhas NEs", len(df_s6['ne'].unique()))
                c2.metric("🟢 Até 3d", len(df_s6[df_s6['dias'] <= 3]))
                c3.metric("🟡 4-7d", len(df_s6[(df_s6['dias'] > 3) & (df_s6['dias'] <= 7)]))
                c4.metric("🔴 >7d", len(df_s6[df_s6['dias'] > 7]))

                st.dataframe(df_s6[['nup', 'ose', 'ne', 'dias']].sort_values(by='dias', ascending=False), use_container_width=True)
                st.divider()
                
                ne_alvo = st.selectbox("Selecione a NE para gerenciar:", [""] + sorted(df_s6['ne'].unique().tolist()), key="fisc_sel_ne_final")
                
                if ne_alvo:
                    df_ne_fisc = df_s6[df_s6['ne'] == ne_alvo].copy()
                    v_total = df_ne_fisc['valor_liquido'].apply(limpar_valor).sum()
                    faturas_txt = ", ".join(df_ne_fisc['Numero_da_fatura'].astype(str).tolist())
                    ose_txt = df_ne_fisc['ose'].iloc[0]
                    
                    st.markdown(f"#### 📝 Gestão da NE: **{ne_alvo}** ({ose_txt})")
                    
                    col_f1, col_f2 = st.columns(2)
                    
                    with col_f1:
                        st.markdown("##### 📤 1. Informar Nota Fiscal")
                        nf_in = st.text_input("Número da NF recebida:", placeholder="Ex: 2026/550", key="in_nf_fisc_final")
                        
                        if st.button("🚀 Salvar NF e Enviar p/ Liquidação", use_container_width=True, key="btn_nf_fisc_final"):
                            if nf_in:
                                with st.spinner("Atualizando processos..."):
                                    for n in df_ne_fisc['nup'].tolist():
                                        cell = aba_p.find(n)
                                        if cell:
                                            # 1. Grava a NF na Coluna P (16)
                                            aba_p.update_cell(cell.row, 16, nf_in) 
                                            # 2. Move para Status 7 (Liquidação)
                                            mover_status(n, 7)
                                            # 3. Registra Log
                                            registrar_acao(n, "N/A", "NF_CADASTRADA_FISCAL", f"NF: {nf_in}")
                                            
                                st.success(f"✅ NF {nf_in} salva! Processo movido para Liquidação.")
                                time.sleep(1.5)
                                st.rerun()
                            else:
                                st.warning("⚠️ Informe o número da NF.")


                    with col_f2:
                        st.markdown("#### 📧 2. Solicitação de Nota Fiscal")
                        
                        # --- BUSCA DE DADOS E CONFIGURAÇÃO DE CONTATOS ---
                        try:
                            # 1. Dados da Tabela-A
                            cnpj_alvo = str(df_ne_fisc['cnpj'].iloc[0]).strip().split('.')[0]
                            df_tabela_a = pd.DataFrame(sh.worksheet(ABA_TABELA_A).get_all_records())
                            
                            linha_ose = df_tabela_a[df_tabela_a['CNPJ'].astype(str).str.contains(cnpj_alvo)]
                            
                            if not linha_ose.empty:
                                info_ose = linha_ose.iloc[0]
                                email_destino = info_ose.get('E-mail Principal da OSE', info_ose.get('Email Principal da OSE', ""))
                                email_gestor_t = info_ose.get('E-mail do Gestor Titular', info_ose.get('Email do Gestor Titular', ""))
                                email_gestor_s = info_ose.get('E-mail do Gestor Substituto', info_ose.get('Email do Gestor Substituto', ""))
                            else:
                                email_destino = "faturamento_ose@gmail.com"
                                email_gestor_t, email_gestor_s = "", ""

                            # 2. Dados do Executor (Logado)
                            df_users = pd.DataFrame(sh.worksheet(ABA_USUARIOS).get_all_records())
                            user_id_atual = str(st.session_state.user_id).strip()
                            match_user = df_users[df_users['NIP'].astype(str).str.strip() == user_id_atual]
                            
                            email_execucao = "hnbra.execucaofinanceira@gmail.com"
                            
                            if not match_user.empty:
                                email_executor = match_user.iloc[0].get('E-mail', match_user.iloc[0].get('Email', email_execucao))
                            else:
                                email_executor = email_execucao
                            
                            # 3. Monta lista de CC
                            lista_cc = list(set([e for e in [email_gestor_t, email_gestor_s, email_executor, email_execucao] if e]))
                            cc_string = ", ".join(lista_cc)
                            
                        except Exception as e:
                            st.error(f"Erro ao processar contatos: {e}")
                            email_destino = "aguardando_dados@ose.com"
                            cc_string = email_execucao

                        # --- CONTEÚDO DO E-MAIL (LIMPO) ---
                        # O .replace('\n', '') aqui evita o erro de 'Header values'
                        assunto_sugerido = f"SOLICITAÇÃO DE NOTA FISCAL - NE {ne_alvo} - {ose_txt}".replace('\n', '').strip()
                        
                        corpo_email = (
                            f"À Gerência de Faturamento da {ose_txt},\n\n"
                            f"Informamos que a Nota de Empenho nº **{ne_alvo}**, no valor total de **R$ {v_total:,.2f}**, "
                            f"referente às faturas **{faturas_txt}**, já encontra-se disponível.\n\n"
                            f"Dessa forma, solicita-se a emissão e o envio da respectiva Nota Fiscal para o e-mail: {email_executor}, "
                            f"conforme trâmite de liquidação e pagamento.\n\n"
                            f"Atenciosamente,\n\n"
                            f"Fiscalização de Contratos - SISAFA-NAVAL"
                        )
                        
                        # --- INTERFACE DE CONFERÊNCIA ---
                        with st.container(border=True):
                            st.write(f"📩 **Para:** {email_destino}")
                            st.write(f"📎 **CC:** {cc_string}")
                            st.divider()
                            # Capturamos os inputs em variáveis para usar no envio
                            assunto_final = st.text_input("Assunto:", value=assunto_sugerido, key="email_sub_fisc_v3")
                            msg_final = st.text_area("Corpo da mensagem:", value=corpo_email, height=250, key="email_body_fisc_v3")
                        
                        # --- BOTÃO DE ENVIO ---
                        if st.button("📧 Disparar Solicitação Oficial", use_container_width=True, key="btn_fisc_send_v3"):
                            if not email_destino or email_destino == "aguardando_dados@ose.com":
                                st.error("⚠️ Erro: E-mail de destino inválido.")
                            else:
                                with st.spinner("Enviando e-mail..."):
                                    # Chamada da função genérica com os dados conferidos/editados
                                    sucesso = enviar_email_generico(
                                        destinatario=email_destino,
                                        assunto=assunto_final,
                                        corpo=msg_final,
                                        cc=lista_cc
                                    )
                                    
                                    if sucesso:
                                        registrar_acao(df_ne_fisc['nup'].iloc[0], "N/A", "EMAIL_SOLICITACAO_NF", f"Para: {email_destino}")
                                        st.success("Solicitação enviada com sucesso!")
                                        time.sleep(1)
                                        st.rerun()
                                    else:
                                        st.error("❌ Falha técnica no servidor de e-mail.")


        # 3. ABA: RELACIONAMENTO
        with tab_rel:
            st.subheader("💬 Central de Relacionamento (Fiscal/OSE)")
            try:
                aba_msg = sh.worksheet(ABA_MENSAGENS)
                df_msg = pd.DataFrame(aba_msg.get_all_records())
                
                if not is_global:
                    cnpjs_meus = df_fiscal['CNPJ'].astype(str).str.split('.').str[0].tolist()
                    df_msg = df_msg[df_msg['cnpj_ose'].astype(str).str.contains('|'.join(cnpjs_meus))].copy()

                if df_msg.empty:
                    st.info("Nenhuma mensagem pendente com suas OSEs.")
                else:
                    st.dataframe(df_msg[['nup', 'cnpj_ose', 'assunto', 'status_resposta']], use_container_width=True)
                    nup_interacao = st.selectbox("Selecione o NUP para responder:", [""] + df_msg['nup'].unique().tolist(), key="sb_rel_fisc")
                    
                    if nup_interacao:
                        dados_m = df_msg[df_msg['nup'] == nup_interacao].iloc[0]
                        st.chat_message("user").write(f"**OSE ({dados_m['cnpj_ose']}):** {dados_m['mensagem_corpo']}")
                        resp_fisc = st.text_area("Sua resposta oficial:", key="resp_fisc_msg")
                        if st.button("📤 Enviar Resposta"):
                            registrar_acao(nup_interacao, "N/A", "RESPOSTA_FISCAL", "Fiscal respondeu via sistema.")
                            st.success("Resposta enviada!")
                            time.sleep(1); st.rerun()
            except Exception as e:
                st.error(f"Erro no módulo de relacionamento: {e}")
    


    elif st.session_state.modulo_ativo == "GERENCIAL" or st.session_state.modulo_ativo == "ADMIN":
        st.header("📈 Dashboard")
        st.metric("Economia", f"R$ {pd.to_numeric(df['glosa']).sum():,.2f}")
        st.bar_chart(df['status'].value_counts())



    elif st.session_state.modulo_ativo == "OSE":
        st.header("🏥 Portal da OSE")
        
        # --- 1. PREPARAÇÃO DOS DADOS (A VACINA DOS 14 DÍGITOS) ---
        user_cnpj = str(st.session_state.user_id).strip().zfill(14)

        # Preparando a Tabela-A (Dados do Fiscal)
        df_tabela_a = carregar_dados_cache(ABA_TABELA_A)
        df_tabela_a.columns = [c.strip().replace(' ', '_').upper() for c in df_tabela_a.columns]
        
        if 'CNPJ' in df_tabela_a.columns:
            # Limpamos o CNPJ da Tabela-A
            df_tabela_a['CNPJ_LIMPO'] = df_tabela_a['CNPJ'].astype(str).str.split('.').str[0].str.strip().str.zfill(14)
            dados_minha_ose = df_tabela_a[df_tabela_a['CNPJ_LIMPO'] == user_cnpj].copy()
        else:
            dados_minha_ose = pd.DataFrame()

        # Preparando as Faturas (df principal)
        if 'cnpj' in df.columns:
            # Limpamos o CNPJ da base de processos
            df['cnpj_limpo'] = df['cnpj'].astype(str).str.split('.').str[0].str.strip().str.zfill(14)
            df_minhas_faturas = df[df['cnpj_limpo'] == user_cnpj].copy()
        else:
            df_minhas_faturas = pd.DataFrame()

        # --- NÃO SEI SE ESTÁ CERTO ---

        mapa_status_fisc = {
        1: "1 - FATURA CADASTRADA", 
        2: "2 - EM AUDITAGEM", 
        3: "3 - AUDITADA",
        4: "4 - AGUARDANDO EMISSÃO DE NE", 
        5: "5 - FATURA EMPENHADA",
        6: "6 - AGUARDANDO EMISSÃO DE NF", 
        7: "7 - EM LIQUIDAÇÃO",
        8: "8 - FATURA LIQUIDADA", 
        9: "9 - FATURA PAGA"
        }


        # --- 2. INTERFACE DAS ABAS ---
        tab_visao, tab_rel = st.tabs(["🔭 Visão Geral", "💬 Relacionamento"])

        # --- 1. ABA: VISÃO GERAL ---
        with tab_visao:
            # Seção: Fiscal do meu Contrato
            st.subheader("👮 Fiscal do meu contrato")
            if not dados_minha_ose.empty:
                # Mapeamento para exibição conforme solicitado
                cols_fiscal = {
                    "NIP_DO_GESTOR_TITULAR": "NIP",
                    "GESTOR_TITULAR": "Nome do Fiscal do Contrato",
                    "GESTOR_SUBSTITUTO": "Fiscal Substituto"
                }
                existentes = [c for c in cols_fiscal.keys() if c in dados_minha_ose.columns]
                st.table(dados_minha_ose[existentes].rename(columns=cols_fiscal))
            else:
                st.info("Informações do fiscal ainda não vinculadas para este CNPJ.")

            st.divider()

            # Seção: Minhas Faturas
            st.subheader("📑 Minhas faturas")
            if df_minhas_faturas.empty:
                st.warning(f"Nenhuma fatura encontrada para o CNPJ: {user_cnpj}")
            else:
                # Mapeamento de Status
                df_minhas_faturas['Situação'] = df_minhas_faturas['status'].map(mapa_status_fisc)
                
                # Seleção e Renomeação conforme seu pedido
                # 'Mes_sigla' e 'Ano de Competência' mapeados de 'mes' e 'ano'
                mapa_colunas_exibicao = {
                    'Numero_da_fatura': 'Nº da fatura',
                    'valor_apresentado': 'Valor Apresentado',
                    'valor_glosa': 'Glosa',
                    'valor_liquido': 'Valor líquido',
                    'mes_sigla': 'Mês de entrada no HNBra',
                    'ano_competencia': 'Ano de Competência',
                    'ne': 'NE',
                    'nf': 'NF',
                    'ob': 'OB',
                    'Situação': 'Situação da Fatura'
                }
                
                colunas_validas = [c for c in mapa_colunas_exibicao.keys() if c in df_minhas_faturas.columns]
                
                st.dataframe(
                    df_minhas_faturas[colunas_validas].rename(columns=mapa_colunas_exibicao).sort_values(by='ano_competencia', ascending=False),
                    use_container_width=True,
                    hide_index=True
                )

        # --- 2. ABA: RELACIONAMENTO ---
        with tab_rel:
            st.subheader("💬 Central de Relacionamento")
            st.markdown("Utilize este espaço para tirar dúvidas sobre faturas específicas.")
            
            if df_minhas_faturas.empty:
                st.info("Você precisa ter faturas cadastradas para iniciar um contato.")
            else:
                # OSE escolhe a fatura pelo NUP ou Número
                opcoes_fatura = df_minhas_faturas['nup'].tolist()
                nup_selecionado = st.selectbox("Sobre qual processo (NUP) você deseja falar?", [""] + opcoes_fatura)
                
                if nup_selecionado:
                    # Busca dados da fatura para mostrar o contexto
                    info_f = df_minhas_faturas[df_minhas_faturas['nup'] == nup_selecionado].iloc[0]
                    st.info(f"Contexto: Fatura {info_f['Numero_da_fatura']} | Status Atual: {info_f['Situação']}")
                    
                    # Área de Mensagem
                    with st.container(border=True):
                        st.write("📤 **Nova Mensagem:**")
                        assunto = st.text_input("Assunto da dúvida:", placeholder="Ex: Prazo de pagamento, Recurso de Glosa...")
                        mensagem = st.text_area("Descreva sua solicitação:")
                        
                        if st.button("Enviar Mensagem Oficial"):
                            if assunto and mensagem:
                                with st.spinner("Enviando..."):
                                    # Aqui usamos a função de registrar ação para que os militares vejam o log
                                    registrar_acao(
                                        nup_selecionado, 
                                        info_f['Numero_da_fatura'], 
                                        "CONTATO_OSE", 
                                        f"Assunto: {assunto} | Msg: {mensagem}"
                                    )
                                    
                                    # Se você tiver uma aba de mensagens específica (ABA_MENSAGENS)
                                    # você pode adicionar um append_row aqui também.
                                    
                                st.success("Sua mensagem foi enviada para a equipe do HNBra!")
                                time.sleep(1)
                                st.rerun()
                            else:
                                st.warning("Por favor, preencha o assunto e a mensagem.")