import streamlit as st
import gspread
from google.oauth2 import service_account 
import os
from PIL import Image
import shutil  
import base64
import pandas as pd
from datetime import datetime
import time
import gspread.exceptions
import smtplib
import re
from email.message import EmailMessage
import plotly.express as px
import pm4py
from pm4py.objects.log.util import dataframe_utils
from pm4py.objects.conversion.log import converter as log_converter
from fpdf import FPDF
import io
import json
import tempfile
import kaleido

# =================================================================
# CONFIGURAÇÃO DO MOTOR GRÁFICO (GRAPHVIZ)
# =================================================================
# Tenta localizar o executável 'dot' automaticamente no sistema
dot_path = shutil.which("dot")

if dot_path:
    # Se achou, forçamos a variável que o PM4Py exige
    os.environ["GRAPHVIZ_DOT"] = dot_path
    dot_dir = os.path.dirname(dot_path)
    if dot_dir not in os.environ["PATH"]:
        os.environ["PATH"] += os.pathsep + dot_dir
else:
    # Caso o sistema não informe o caminho, tentamos os padrões do Linux
    caminhos_manuais = ["/usr/bin/dot", "/usr/local/bin/dot"]
    for caminho in caminhos_manuais:
        if os.path.exists(caminho):
            os.environ["GRAPHVIZ_DOT"] = caminho
            os.environ["PATH"] += os.pathsep + os.path.dirname(caminho)
            break

# --- python -m streamlit run app.py ---
# --- CONFIGURAÇÕES E CAMINHOS ---
ID_PLANILHA = "1NS9zdzNFcHjQ7zFpEysuU-udrrV1VaM7nPY7LjHk3Qk"
ABA_USUARIOS = "SISAFA-NAVAL-Usuarios"
ABA_PROCESSOS = "SISAFA-NAVAL-processos"
ABA_LOGS_ACOES = "SISAFA-NAVAL-logs_acoes"
ABA_HISTORICO = "SISAFA-NAVAL-historico"
ABA_TABELA_A = "SISAFA-NAVAL-Tabela-A"
ABA_MENSAGENS = "SISAFA-NAVAL-mensagens"
ABA_AUDITORIA = "SISAFA-NAVAL-Auditoria"
ABA_AUDITORIA_GLOSA = "SISAFA-NAVAL-Auditoria-glosa"
ABA_TABELA_GLOSA = "SISAFA-NAVAL-Tabela-de-referencia-de-glosa"
ABA_RASCUNHO = "SISAFA-NAVAL-Rascunhos"

# Localiza a pasta do projeto
pasta_projeto = os.path.dirname(os.path.abspath(__file__))
caminho_logo = os.path.join(pasta_projeto, "LOGO-SISAFA-NAVAL.png")
caminho_logo_relatorio = os.path.join(pasta_projeto, "SISAFA-NAVAL-relatorio.png")
caminho_mascote = os.path.join(pasta_projeto, "canto_inferior_direito_da_tela_de_apresentacao.png")
caminho_mapeamento = os.path.join(pasta_projeto, "mapeamento-de-processo.png")
caminho_favicon = os.path.join(pasta_projeto, "Favicon-SISAFA-NAVAL.png")
caminho_escudo_dsm = os.path.join(pasta_projeto, "Simbolo-DSM_SISAFA.png")

icone = Image.open(caminho_favicon)

st.set_page_config(
    page_title="SISAFA NAVAL ⚓", 
    layout="centered", 
    page_icon=icone,
    initial_sidebar_state="expanded" 
)

# --- ESTILIZAÇÃO CSS  ---

st.markdown("""
    <style>
        /* 1. O TIRO DE SNIPER NA SETINHA DA BARRA LATERAL (<<) */
        [data-testid="stSidebarHeader"] button,
        [data-testid="stSidebarCollapseButton"] {
            display: none !important;
            pointer-events: none !important;
        }

        /* --- MANTENHA O RESTO DO SEU CÓDIGO AQUI --- */
        /* Exemplo do que já tínhamos: */
        header[data-testid="stHeader"] { background-color: transparent !important; }
        .viewerBadge_container, [class^="viewerBadge"], [data-testid="stToolbar"], .stAppDeployButton { display: none !important; }
        #MainMenu { display: none !important; }
        footer { display: none !important; }
        
        div.stButton > button {
            background-color: #2e6b54 !important;
            color: white !important;
            border-radius: 8px !important;
            border: none !important;
            font-weight: bold !important;
        }
        div.stButton > button:hover { background-color: #1e4536 !important; }
        .block-container { padding-top: 2rem !important; }
    </style>
""", unsafe_allow_html=True)


# =================================================================
# --- CONEXÃO DIRETA E OTIMIZADA (FOCO EM PERFORMANCE E COTA) ---
# =================================================================

@st.cache_resource(ttl=3600)
def obter_cliente_google():
    """
    Autentica no Google de forma direta e guarda o cliente em cache por 1 hora.
    Evita múltiplas requisições e protege contra o Erro 429.
    """
    try:
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        
        if "gcp_service_account" not in st.secrets:
            st.error("🚨 Chaves do Google (gcp_service_account) não encontradas no st.secrets!")
            return None
            
        creds_info = st.secrets["gcp_service_account"]
        
        if isinstance(creds_info, str):
            creds_info = json.loads(creds_info.strip())
        
        if "private_key" in creds_info:
            creds_info["private_key"] = creds_info["private_key"].replace("\\n", "\n").strip()
        
        creds = service_account.Credentials.from_service_account_info(creds_info, scopes=scope)
        return gspread.authorize(creds)
        
    except Exception as e:
        st.error(f"❌ Erro na autenticação com o Google: {e}")
        return None

def abrir_planilha_mestre():
    """Garante a ligação à planilha mestre utilizando o cliente em cache de forma segura"""
    if 'spreadsheet_objeto' not in st.session_state or st.session_state.spreadsheet_objeto is None:
        try:
            client = obter_cliente_google()
            if client is None:
                return None
            st.session_state.spreadsheet_objeto = client.open_by_key(ID_PLANILHA)
        except Exception as e:
            st.error(f"Erro crítico ao aceder à planilha mestre: {e}")
            return None
            
    return st.session_state.spreadsheet_objeto

# =================================================================
# --- INICIALIZAÇÃO GLOBAL LIMPA ---
# =================================================================
sh = abrir_planilha_mestre()

if sh is not None:
    try:
        # Define a aba de processos globalmente
        aba_p = sh.worksheet(ABA_PROCESSOS)
    except Exception as e:
        st.error(f"Erro ao aceder à aba de processos: {e}")
else:
    st.warning("⚠️ Sistema temporariamente indisponível. Por favor, atualize a página (F5) dentro de alguns minutos.")




def registrar_historico(nup, fatura, origem, destino, valor, obs=""):
    try:
        # Usa o 'sh' global, sem fazer novo login!
        if sh:
            aba = sh.worksheet(ABA_HISTORICO)
            agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            aba.append_row([agora, nup, fatura, origem, destino, st.session_state.user_id, valor, obs])
    except Exception as e:
        st.error(f"Erro na aba HISTÓRICO: {e}")

def registrar_acao(nup, fatura, acao, detalhes=""):
    try:
        # Usa o 'sh' global
        if sh:
            aba = sh.worksheet(ABA_LOGS_ACOES)
            agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            aba.append_row([str(datetime.now().timestamp()), nup, fatura, acao, st.session_state.user_id, agora, detalhes])
    except: 
        pass

def mover_status(nup, novo_status, auditor_nip=None, obs_texto=None, valor_glosa=None, valor_liq=None):
    if not sh:
        return False
        
    aba_p = sh.worksheet(ABA_PROCESSOS)
    cell = aba_p.find(nup)
    
    if cell:
        dados_atuais = aba_p.row_values(cell.row)
        status_origem = dados_atuais[10] 
        fatura = dados_atuais[4]
        valor_atual = valor_liq if valor_liq is not None else dados_atuais[7]
        
        agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Atualizações mantidas, mas sem gerar novo login
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
            <p><b>Justificativa resumida:</b> {justificativa}</p>
            <p>O relatório de glosa seguirá formalmente assim que possível, para o caso de necessidade de interposição de recurso pela via administrativa.</p>
            <br>
            <p>Cordialmente,</p>
            <p><b>Equipe de Auditoria em Saúde</b><br>
            <p><b>☎️ Telefone: 3445-7318 ☎️</b><br>
            Sistema de Acompanhamento de Faturas do Hospital Naval de Brasília</p>
            <hr>
            <p><small style="color: gray;">🛑⚠️ E-mail automático gerado pelo Sistema de Acompanhamento de Faturas do Hospital Naval de Brasília. Favor não responder.⚠️🛑</small></p>
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

def tratar_texto_pdf(texto):
    if not texto:
        return ""
    # Remove caracteres que o PDF não entende (emojis, aspas especiais, etc)
    return str(texto).encode('latin-1', 'ignore').decode('latin-1')

def gerar_relatorio_pdf(dados_nup, auditor_nome, total_glosa, justificativa, valores_dict, listas_grupos, lista_g6, v_apres):
    from fpdf import FPDF
    import datetime

    def limpar(txt):
        if not txt: return ""
        return str(txt).encode('latin-1', 'ignore').decode('latin-1')

    # (1) VERTICAL E LIMPO: Segue exatamente o modelo da imagem
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    # --- 0. MARCA D'ÁGUA (Fundo do documento) ---
    try:
        # Centralizada como no seu modelo
        pdf.image('SISAFA-NAVAL-relatorio.png', x=60, y=95, w=90)
    except:
        pass 

    # --- 1. CABEÇALHO ---
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 6, "Hospital Naval de Brasília (HNBra)", ln=True, align='C')
    pdf.cell(0, 6, limpar("Relatório de Auditoria de Fatura"), ln=True, align='C')
    
    # --- 2. DADOS DO PROCESSO ---
    pdf.ln(5)
    pdf.set_font("Arial", 'B', 10)
    pdf.set_fill_color(230, 230, 230)
    pdf.cell(0, 7, "DADOS DO PROCESSO", 1, ln=True, fill=True)
    
    pdf.set_font("Arial", '', 10)
    pdf.cell(0, 7, limpar(f"NUP: {dados_nup.get('nup', 'N/A')} | Fatura: {dados_nup.get('Numero_da_fatura', 'N/A')}"), 1, ln=True)
    pdf.cell(0, 7, limpar(f"Auditor(a): {auditor_nome} | Data/Hora: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"), 1, ln=True)

    # --- 3. RESUMO FINANCEIRO DA AUDITORIA ---
    pdf.ln(5)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(0, 7, "RESUMO FINANCEIRO DA AUDITORIA", 1, ln=True, fill=True)
    
    pdf.set_font("Arial", '', 10)
    v_liq = v_apres - total_glosa
    pdf.cell(0, 7, f"Valor Apresentado: R$ {v_apres:,.2f}", 1, ln=True)
    pdf.cell(0, 7, f"(-) Valor da Glosa: R$ {total_glosa:,.2f}", 1, ln=True)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(0, 7, f"(=) Valor Liquido: R$ {v_liq:,.2f}", 1, ln=True)

    # --- 4. SEÇÃO DE GLOSA ---
    pdf.ln(5)
    houve_glosa = "SIM" if total_glosa > 0 else "NÃO"
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(0, 7, f"Houve Glosa: {houve_glosa}", 1, ln=True)
    
    pdf.cell(0, 6, "Justificativa Técnica:", "LR", ln=True)
    pdf.set_font("Arial", '', 9)
    pdf.multi_cell(0, 5, limpar(justificativa if justificativa else "N/A"), "LRB")

    # --- 5. TABELA DE ITENS (Onde aparece o SIAD, etc) ---
    pdf.ln(5)
    pdf.set_font("Arial", 'B', 9)
    pdf.cell(140, 7, limpar("Descrição do Procedimento/Exame"), 1, 0, 'L', True)
    pdf.cell(50, 7, limpar("Valor (R$)"), 1, 1, 'C', True)
    
    pdf.set_font("Arial", '', 9)
    # Lógica para mostrar apenas os campos que têm valor
    for grupo in listas_grupos:
        for campo in grupo:
            val = float(valores_dict.get(campo, 0))
            if val > 0:
                pdf.cell(140, 7, limpar(campo), 1, 0, 'L')
                pdf.cell(50, 7, f"{val:,.2f}", 1, 1, 'R')

    # Adiciona itens extras do Grupo VI
    for item in lista_g6:
        if item['tipo'] and float(item.get('valor', 0)) > 0:
            pdf.cell(140, 7, limpar(f"OUTROS: {item['tipo']} - {item['desc']}"), 1, 0, 'L')
            pdf.cell(50, 7, f"{float(item['valor']):,.2f}", 1, 1, 'R')

    # --- 6. TOTAL FINAL (RODAPÉ DA TABELA) ---
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(140, 8, limpar("TOTAL FINAL AUDITADO (VALOR LÍQUIDO)"), 1, 0, 'L', True)
    pdf.cell(50, 8, f"R$ {v_liq:,.2f}", 1, 1, 'R', True)

    return pdf.output(dest='S').encode('latin-1', 'ignore')


def gerar_relatorio_glosa_pdf(dados_nup, dados_ose, lista_glosas, auditor_info, num_relatorio, justificativa):
    from fpdf import FPDF
    import datetime
    
    def limpar(txt):
        if not txt: return ""
        return str(txt).encode('latin-1', 'ignore').decode('latin-1')

    # (1) PAPEL NA HORIZONTAL: orientation='L' (Landscape)
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    
    # Reduzida a margem inferior para dar mais espaço útil à tabela antes da quebra
    pdf.set_auto_page_break(auto=True, margin=10)

    # --- 0. MARCA D'ÁGUA ---
    try:
        pdf.image('SISAFA-NAVAL-relatorio.png', x=100, y=60, w=100)
    except:
        pass 
    
    # --- 1. CABEÇALHO (EM VERMELHO) ---
    pdf.set_font("Arial", 'B', 9)
    pdf.set_text_color(220, 0, 0) # Vermelho
    pdf.cell(0, 4, limpar("INFORMAÇÃO PESSOAL - ACESSO RESTRITO"), ln=True, align='C')
    pdf.set_font("Arial", '', 8)
    pdf.multi_cell(0, 3, limpar("Art. 5º, Inciso X da Constituição Federal do Brasil/1988\nArt. 31 da Lei nº 12.527/2011 | Art. 55 e Art. 62 do Dec 7.724/2012"), align='C')
    pdf.set_text_color(0, 0, 0) # Reset para Preto

    # --- 2. IDENTIFICAÇÃO INSTITUCIONAL ---
    pdf.ln(4)
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 6, limpar("MARINHA DO BRASIL"), ln=True, align='C')
    pdf.cell(0, 6, limpar("HOSPITAL NAVAL DE BRASÍLIA"), ln=True, align='C')
    pdf.cell(0, 6, limpar("AUDITORIA EM SAÚDE"), ln=True, align='C')
    pdf.set_fill_color(230, 230, 230)
    pdf.cell(0, 10, limpar("RELATÓRIO DE GLOSA"), 1, ln=True, align='C', fill=True)

    # --- 3. DADOS DA OSE ---
    pdf.ln(3)
    pdf.set_font("Arial", 'B', 8)
    pdf.cell(140, 5, limpar("Organização Civil de Saúde (OCS) credenciada"), 1, 0, 'L', True)
    pdf.cell(137, 5, limpar("Nº do Edital de Credenciamento"), 1, 1, 'L', True)
    
    pdf.set_font("Arial", '', 8)
    pdf.cell(140, 6, limpar(dados_ose.get('Razão Social', 'N/A')), 1)
    pdf.cell(137, 6, limpar(dados_ose.get('Numero_edital', 'N/A')), 1, 1)
    
    pdf.set_font("Arial", 'B', 8)
    pdf.cell(70, 5, limpar("Nº do Termo de Credenciamento"), 1, 0, 'L', True)
    pdf.cell(50, 5, limpar("Validade Edital"), 1, 0, 'L', True)
    pdf.cell(157, 5, limpar("Endereço eletrônico da OCS"), 1, 1, 'L', True)
    
    pdf.set_font("Arial", '', 7)
    pdf.cell(70, 6, limpar(dados_ose.get('Termo de credenciamento', 'N/A')), 1)
    pdf.cell(50, 6, limpar(dados_ose.get('Validade_edital', 'N/A')), 1)
    pdf.cell(157, 6, limpar(dados_ose.get('E-mail Principal da OSE', 'N/A')[:100]), 1, 1)

    # --- 4. DADOS DO PROCESSO ---
    pdf.ln(3)
    pdf.set_font("Arial", 'B', 8)
    pdf.cell(60, 5, "NUP", 1, 0, 'L', True)
    pdf.cell(35, 5, limpar("Nº Relatório"), 1, 0, 'L', True)
    pdf.cell(45, 5, limpar("Nº Fatura/Remessa"), 1, 0, 'L', True)
    
    # Novos Campos Financeiros
    pdf.cell(45, 5, limpar("Valor Apresentado"), 1, 0, 'L', True)
    pdf.cell(45, 5, limpar("Glosa Total"), 1, 0, 'L', True)
    pdf.cell(47, 5, limpar("Valor Líquido"), 1, 1, 'L', True)

    # Cálculos
    v_apres_limpo = limpar_valor(dados_nup['valor_apresentado'])
    total_glosa = sum(float(g['valor']) for g in lista_glosas)
    valor_liquido = v_apres_limpo - total_glosa
    ano_atual = datetime.datetime.now().strftime('%y')

    pdf.set_font("Arial", '', 8)
    pdf.cell(60, 7, limpar(dados_nup['nup']), 1)
    pdf.cell(35, 7, f"{num_relatorio}/{ano_atual}", 1)
    pdf.cell(45, 7, limpar(dados_nup['Numero_da_fatura']), 1)
    
    # Valores preenchidos
    pdf.cell(45, 7, f"R$ {v_apres_limpo:,.2f}", 1)
    pdf.set_text_color(200, 0, 0) # Vermelho para a glosa
    pdf.cell(45, 7, f"R$ {total_glosa:,.2f}", 1)
    pdf.set_text_color(0, 0, 0) # Volta para preto
    pdf.set_font("Arial", 'B', 8)
    pdf.cell(47, 7, f"R$ {valor_liquido:,.2f}", 1, 1)

    # --- 5. TABELA DE PACIENTES E GLOSAS ---
    pdf.ln(3)
    pdf.set_font("Arial", 'B', 8)
    pdf.cell(10, 6, "N", 1, 0, 'C', True)
    pdf.cell(50, 6, "Paciente", 1, 0, 'L', True)
    pdf.cell(35, 6, "Valor da Glosa", 1, 0, 'C', True)
    pdf.cell(30, 6, "Tipo", 1, 0, 'C', True)
    pdf.cell(20, 6, "Cod.", 1, 0, 'C', True)
    pdf.cell(132, 6, limpar("Descrição"), 1, 1, 'C', True)

    pdf.set_font("Arial", '', 7)
    
    # 5.1 Correção da Altura Dinâmica da Linha (Evita quebra de célula torta)
    for i, g in enumerate(lista_glosas, 1):
        # Captura as coordenadas iniciais da linha
        start_y = pdf.get_y()
        start_x = pdf.get_x()
        
        # Pinta a célula MultiCell (Descrição) primeiro, para sabermos a altura que ela vai tomar
        pdf.set_xy(start_x + 145, start_y) # 10+50+35+30+20 = 145 de deslocamento X
        pdf.multi_cell(132, 8, limpar(g['just']), 1, 'L')
        
        # A altura real que a linha toda deve ter é baseada no que a multi_cell desenhou
        end_y = pdf.get_y()
        altura_linha = end_y - start_y
        
        # Desenha as células normais (Single Cells) com a altura descoberta
        pdf.set_xy(start_x, start_y)
        pdf.cell(10, altura_linha, f"{i}", 1, 0, 'C')
        pdf.cell(50, altura_linha, limpar(g['paciente']), 1, 0, 'L')
        pdf.cell(35, altura_linha, f"R$ {float(g['valor']):,.2f}", 1, 0, 'R')
        pdf.cell(30, altura_linha, limpar(g['tipo']), 1, 0, 'C')
        pdf.cell(20, altura_linha, limpar(g['cod']), 1, 0, 'C')
        
        # Reposiciona o cursor no final da linha que acabamos de desenhar
        pdf.set_xy(start_x, end_y)

    # --- NOVO CAMPO: OBSERVAÇÃO (Justificativa Técnica) ---
    pdf.ln(2)
    pdf.set_font("Arial", 'B', 9)
    pdf.cell(0, 6, limpar("Observação:"), 0, 1)
    pdf.set_font("Arial", '', 8)
    pdf.multi_cell(0, 5, limpar(justificativa), 1)

    # --- 6. CONTROLE ESPACIAL (PREVENÇÃO DE DESALINHAMENTO) ---
    # Se faltar menos de 45mm para o final da página, força uma página nova
    # Assim, o rodapé e as assinaturas nunca ficam cortados ou flutuando.
    espaco_necessario_assinaturas = 45 
    espaco_restante = 210 - pdf.get_y() # 210mm é a altura de A4 Landscape
    
    if espaco_restante < espaco_necessario_assinaturas:
        pdf.add_page()
        # Reposiciona a marca d'água na nova página, se existir
        try:
            pdf.image('SISAFA-NAVAL-relatorio.png', x=100, y=60, w=100)
        except:
            pass

    # --- 7. RODAPÉ E DATAS ---
    pdf.ln(5)
    pdf.set_font("Arial", 'B', 8)
    pdf.cell(140, 5, limpar("Legendas: Adm: Administrativa | Téc: Técnica"), 0, 0, 'L')
    pdf.cell(137, 5, limpar(f"Brasília, na data de assinatura."), 0, 1, 'R')
    
    # --- 8. ASSINATURAS LADO A LADO ---
    pdf.ln(10)
    largura_col = 92
    y_assinaturas = pdf.get_y()
    pdf.set_font("Arial", 'B', 8)

    # Coluna 1: GISELLE
    pdf.set_xy(10, y_assinaturas)
    pdf.multi_cell(largura_col, 4, limpar("GISELLE BITENCOURT\nCapitão de Fragata (S)\nEncarregada da Auditoria em Saúde"), 0, 'C')

    # Coluna 2: CAMILA
    pdf.set_xy(10 + largura_col, y_assinaturas)
    pdf.multi_cell(largura_col, 4, limpar("CAMILA GUERRA FELICIANO MORAIS\nCapitão-Tenente (S)\nEncarregada da Auditoria de Contas Hospitalares"), 0, 'C')

    # Coluna 3: AUDITOR
    pdf.set_xy(10 + (largura_col * 2), y_assinaturas)
    pdf.multi_cell(largura_col, 4, limpar(f"{auditor_info['nome']}\nAuditor Responsável"), 0, 'C')

    # --- 9. RODAPÉ FINAL (FIXADO NA ÚLTIMA FOLHA) ---
    # Forçamos a posição para o rodapé ficar estático no final da página
    pdf.set_y(-25) 
    pdf.set_text_color(220, 0, 0)
    pdf.set_font("Arial", 'B', 8)
    pdf.cell(0, 4, limpar("INFORMAÇÃO PESSOAL - ACESSO RESTRITO"), ln=True, align='C')
    pdf.set_font("Arial", '', 8)
    rodape_legal = (
        "Art. 5º, Inciso X, da Constituição Federal do Brasil/1988\n"
        "Art. 31 da Lei nº 12.527, de 18 de novembro de 2011\n"
        "Art. 55 ao Art. 62 do Dec nº 7.724, de 16 de maio de 2012"
    )
    pdf.multi_cell(0, 3, limpar(rodape_legal), align='C')
    pdf.set_text_color(0, 0, 0)

    return pdf.output(dest='S').encode('latin-1', 'ignore')

def obter_proximo_numero_relatorio(sh):
    try:
        aba_glosa = sh.worksheet("SISAFA-NAVAL-Auditoria-glosa")
        # Pega todos os valores da coluna "Numero_relatorio_glosa" (Coluna L / índice 11)
        coluna_relatorios = aba_glosa.col_values(12) 
        
        # Remove o cabeçalho e filtra apenas números
        numeros = []
        for val in coluna_relatorios[1:]:
            try:
                numeros.append(int(float(val)))
            except:
                continue
        
        if not numeros:
            return 1
        return max(numeros) + 1
    except Exception as e:
        st.error(f"Erro ao gerar numeração automática: {e}")
        return 1 # Fallback para não travar o sistema


def obter_proximo_numero_glosa():
    max_tentativas = 3
    
    # Se a planilha não carregou por algum motivo, encerra para evitar erro
    if sh is None:
        st.error("Erro de conexão com a planilha mestre.")
        return None
        
    for tentativa in range(max_tentativas):
        try:
            # Tenta acessar a aba de glosa
            aba_glosa_folha = sh.worksheet("SISAFA-NAVAL-Auditoria-glosa")
            
            # --- SUBSTITUA PELO SEU CÓDIGO DE CONTAGEM ---
            # Exemplo de como pegar o número da próxima linha:
            valores_coluna = aba_glosa_folha.col_values(1)
            quantidade_registros = len(valores_coluna) - 1 # Desconta o cabeçalho
            proximo_numero = quantidade_registros + 1
            
            return proximo_numero
            
        except gspread.exceptions.APIError as e:
            # Se o Google reclamar que acessou rápido demais (Erro 429)
            if e.response.status_code == 429:
                if tentativa < max_tentativas - 1:
                    # O SISAFA respira por alguns segundos e tenta de novo sozinho
                    time.sleep((tentativa + 1) * 2) 
                else:
                    st.error("Servidores do Google ocupados. Aguarde alguns segundos e tente novamente.")
                    st.stop()
            else:
                # Mostra outros erros (como aba com nome errado)
                raise e


import tempfile
import os
import matplotlib.pyplot as plt
from fpdf import FPDF

def gerar_relatorio_ose_pdf(ose_nome, df_ose, volume_total, qtd_total, fig_pie):
    from fpdf import FPDF
    import tempfile
    import os
    import matplotlib.pyplot as plt

    # Função limpar apenas para garantir que não trave com acentos imprevistos, 
    # mas sem aquela substituição manual feia de emojis.
    def limpar(txt):
        if not txt: return ""
        return str(txt).encode('latin-1', 'ignore').decode('latin-1')

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_margins(10, 10, 10)

    # --- 1. CABEÇALHO ---
    pdf.set_xy(10, 10)
    pdf.set_font("Arial", 'B', 14)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(190, 8, "HOSPITAL NAVAL DE BRASÍLIA", 0, 1, 'C')
    pdf.cell(190, 8, "RELATÓRIO DE SITUAÇÃO DAS FATURAS", 0, 1, 'C')
    
    # Destaque do nome da OSE na cor Verde SISAFA
    pdf.set_font("Arial", 'B', 12)
    pdf.set_text_color(0, 230, 118) # Verde Neon
    pdf.cell(190, 8, limpar(ose_nome.upper()), 0, 1, 'C')
    pdf.set_text_color(0, 0, 0) # Volta para preto
    pdf.set_font("Arial", 'I', 10)
    pdf.cell(190, 8, "SISAFA NAVAL", 0, 1, 'C')
    pdf.ln(5)

    # --- 2. PAINEL FINANCEIRO ESTILIZADO ---
    pdf.set_y(52)
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(190, 8, "Painel Financeiro", 0, 1, 'L')
    
    # Fundo mais escuro para simular os "Cards" do sistema
    pdf.set_fill_color(40, 40, 40) # Cinza Escuro (Dark Mode)
    pdf.rect(10, 62, 90, 22, 'F') 
    pdf.rect(110, 62, 90, 22, 'F')
    
    # Textos dos Cards em Branco e Verde Neon
    pdf.set_text_color(255, 255, 255) # Branco
    pdf.set_font("Arial", '', 9)
    pdf.text(15, 70, "Volume Total")
    pdf.text(115, 70, "Total Faturas")
    
    pdf.set_text_color(0, 230, 118) # Verde Neon
    pdf.set_font("Arial", 'B', 14)
    pdf.text(15, 79, f"R$ {volume_total:,.2f}")
    pdf.text(115, 79, f"{qtd_total} unidades")
    pdf.set_text_color(0, 0, 0) # Reset para preto para o resto da página

    # --- 3. GRÁFICO (EFEITO NEON/MODERNO) ---
    if fig_pie:
        try:
            json_data = fig_pie.data[0].to_plotly_json()
            labels = json_data.get('names', json_data.get('labels', []))
            values = json_data.get('values', [])
            
            # Paleta de Cores Estilo "Neon Dashboard"
            cores_neon = [
                '#00E676', '#2979FF', '#FF1744', '#FFEA00', '#D500F9', 
                '#00B8D4', '#FF9100', '#76FF03', '#F50057'
            ]
            
            # Criação do gráfico com fundo transparente e fatias separadas (explode)
            plt.figure(figsize=(6, 3), dpi=150)
            
            # O "explode" separa um pouco as fatias, dando um ar mais tecnológico
            separacao = [0.03] * len(values) 
            
            wedges, texts, autotexts = plt.pie(
                values, labels=labels, autopct='%1.1f%%', 
                colors=cores_neon[:len(values)], explode=separacao,
                textprops={'fontsize': 8, 'fontweight': 'bold', 'color': '#333333'},
                wedgeprops={'linewidth': 1, 'edgecolor': 'white'}
            )
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                plt.savefig(tmp.name, bbox_inches='tight', transparent=True)
                plt.close()
                pdf.image(tmp.name, x=30, y=90, w=120)
                os.remove(tmp.name)
        except: pass

    # --- 4. TABELA ---
    pdf.set_y(155)
    pdf.set_fill_color(0, 230, 118) # Cabeçalho da tabela em Verde Neon
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", 'B', 9)
    pdf.cell(40, 8, "Nº Fatura", 1, 0, 'C', True)
    pdf.cell(40, 8, "Valor (R$)", 1, 0, 'C', True)
    pdf.cell(110, 8, "Situação", 1, 1, 'C', True)

    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", '', 8)
    for _, row in df_ose.sort_values(by='status').iterrows():
        if pdf.get_y() > 245:
            pdf.add_page()
            pdf.set_y(10)
        
        pdf.cell(40, 7, limpar(str(row.get('Numero_da_fatura', 'SN'))), 1, 0, 'C')
        pdf.cell(40, 7, f"{float(row.get('v_liq_num', 0)):,.2f}", 1, 0, 'R')
        pdf.cell(110, 7, limpar(str(row.get('etapa_nome', 'Indefinida'))), 1, 1, 'L')

    # --- 5. RODAPÉ ---
    pdf.set_y(180) 
    if os.path.exists("mapeamento-de-processo.png"):
        pdf.image("mapeamento-de-processo.png", x=10, y=180, w=70, h=0)
    pdf.set_xy(60, 230)
    pdf.set_font("Arial", 'I', 12)
    msg = "Esperamos fortalecer a confiança mútua e\na parceria com o Hospital Naval de Brasília.\nSomos gratos pelo apoio e pela distinta cooperação.\n\nHospital Naval de Brasília - A Saúde Naval no Planalto Central!"
    pdf.multi_cell(150, 6, limpar(msg), 0, 'R')

    return pdf.output(dest='S').encode('latin-1', 'ignore')
    
# --- CONEXÃO GLOBAL BLINDADA (Substitua no topo do arquivo) ---
def obter_sh():
    """Garante a ligação ao Google Sheets com verificação de segurança contra quebras"""
    if 'spreadsheet_objeto' not in st.session_state or st.session_state.spreadsheet_objeto is None:
        try:
            client = obter_cliente_google()
            
            # VALIDAÇÃO DE SEGURANÇA: Se a internet caiu totalmente após as 4 tentativas
            if client is None:
                return None
                
            st.session_state.spreadsheet_objeto = client.open_by_key(ID_PLANILHA)
        except Exception as e:
            # Captura erros específicos como ID inválido ou falta de permissão
            st.error(f"Erro crítico ao aceder à planilha mestre: {e}")
            return None
            
    return st.session_state.spreadsheet_objeto

def salvar_rascunho_auditoria(nup, dados_glosa, valores_cc, justificativa, dados_g6=None):
    """Grava o estado atual da auditagem com logs de depuração"""
    # Garante que, se vier vazio, seja uma lista
    if dados_g6 is None:
        dados_g6 = []
        
    try:
        # 1. Tenta obter o objeto da planilha
        sh_obj = obter_sh()
        if sh_obj is None:
            st.error("❌ Erro de Conexão: O objeto 'sh' está vazio.")
            return False
            
        # 2. Verifica se a aba existe
        try:
            aba_rascunho = sh_obj.worksheet(ABA_RASCUNHO)
        except Exception as e:
            st.error(f"❌ Aba '{ABA_RASCUNHO}' não encontrada na planilha!")
            return False

        # 3. Prepara o pacote de dados (AGORA COM O GRUPO VI)
        import json
        pacote = {
            "glosas": dados_glosa,
            "centro_custo": valores_cc,
            "justificativa": justificativa,
            "grupo6": dados_g6 
        }
        json_dados = json.dumps(pacote)
        agora = datetime.now().strftime("%d/%m/%Y %H:%M")
        
        # 4. Procura o NUP para decidir se ATUALIZA ou CRIA NOVO
        try:
            celula = aba_rascunho.find(str(nup))
            if celula:
                aba_rascunho.update_cell(celula.row, 2, json_dados)
                aba_rascunho.update_cell(celula.row, 3, agora)
                st.toast(f"✅ Rascunho do NUP {nup} atualizado!", icon="🔄")
            else:
                aba_rascunho.append_row([str(nup), json_dados, agora])
                st.toast(f"✅ Novo rascunho criado para o NUP {nup}!", icon="💾")
            return True
        except Exception as e:
            st.error(f"❌ Erro ao escrever na aba: {e}")
            return False

    except Exception as e:
        st.error(f"❌ Erro geral na função de rascunho: {e}")
        return False

def carregar_rascunho(nup):
    """Busca se existe trabalho salvo para este NUP"""
    try:
        sh_obj = obter_sh()
        aba_rasc = sh_obj.worksheet(ABA_RASCUNHO)
        celula = aba_rasc.find(str(nup))
        if celula:
            json_string = aba_rasc.cell(celula.row, 2).value
            return json.loads(json_string)
    except:
        return None


@st.cache_data(ttl=3600)
def obter_tabela_referencia_glosa():
    """Lê a tabela de referência do Google apenas 1 vez por hora."""
    if sh is None: 
        return {}
    try:
        aba_ref = sh.worksheet("SISAFA-NAVAL-Tabela-de-referencia-de-glosa")
        dados_glosa_brutos = aba_ref.get_all_records()
        return {str(row['Cod_glosa']): row['Desc_glosa'] for row in dados_glosa_brutos}
    except Exception as e:
        return {}

# --- CONFIGURAÇÕES DE IMAGEM SEGURAS ---
pasta_projeto = os.path.dirname(os.path.abspath(__file__))
caminho_logo = os.path.join(pasta_projeto, "LOGO-SISAFA-NAVAL.png")
caminho_mascote = os.path.join(pasta_projeto, "canto_inferior_direito_da_tela_de_apresentacao.png")

# Função para carregar imagem sem quebrar o app
def carregar_imagem(caminho):
    return caminho if os.path.exists(caminho) else None

# --- 2. CONEXÃO GLOBAL E DEFINIÇÃO DE 'sh' ---

# 1. LISTA MESTRE DE MUNIÇÃO (Armamento completo para não travar mais)
COLUNAS_MESTRE = [
    'status', 'mes_competencia', 'ano_competencia', 'nup', 
    'valor_apresentado', 'cnpj', 'glosa', 'paciente', 'just',
    'ose', 'v_ap_num', 'data_entrada', 'Numero_da_fatura'
]

# ==========================================
# ENCAIXE DA FUNÇÃO DE LEITURA AQUI
# ==========================================
@st.cache_data(ttl=70)
def carregar_dados_cache(nome_aba):
    """Lê a planilha e já entrega o DF blindado com todas as colunas mestre."""
    try:
        sh_c = abrir_planilha_mestre() 
        
        if sh_c:
            aba = sh_c.worksheet(nome_aba)
            dados = aba.get_all_records()
            
            df = pd.DataFrame(dados) if dados else pd.DataFrame(columns=COLUNAS_MESTRE)
            
            # MANOBRA DE RESTAURAÇÃO INTEGRADA
            for col in COLUNAS_MESTRE:
                if col not in df.columns:
                    df[col] = 0 if col in ['status', 'glosa', 'v_ap_num', 'valor_apresentado'] else ""
            return df
    except Exception as e:
        # Mostra o erro real na tela para podermos diagnosticar
        st.error(f"🚨 ERRO NA CONEXÃO COM O GOOGLE (Aba {nome_aba}): {e}")
    
    return pd.DataFrame(columns=COLUNAS_MESTRE)
# ==========================================

# --- CONEXÃO GLOBAL BLINDADA E CARREGAMENTO ---
try:
    sh = obter_sh() 
    
    if sh is not None:
        # Se a ligação estiver ativa, carrega os dados normalmente
        df = carregar_dados_cache(ABA_PROCESSOS)
    else:
        # SE A INTERNET CAIR TOTALMENTE: O sistema não crasha.
        # Ativa o modo de emergência com uma tabela vazia estruturada para os rascunhos funcionarem
        st.sidebar.warning("⚠️ Modo de Emergência Ativo: Sem ligação ao banco de dados.")
        sh = None
        df = pd.DataFrame(columns=COLUNAS_MESTRE) 

except Exception as e:
    st.sidebar.error("⚠️ Falha grave na inicialização do sistema.")
    sh = None
    df = pd.DataFrame(columns=COLUNAS_MESTRE)

# --- CONTROLE DE SESSÃO ---
if 'logged_in' not in st.session_state: 
    st.session_state.logged_in = False
if 'modulo_ativo' not in st.session_state: 
    st.session_state.modulo_ativo = None
if 'confirmar_secom' not in st.session_state: 
    st.session_state.confirmar_secom = False
if 'confirmar_recebimento' not in st.session_state: 
    st.session_state.confirmar_recebimento = False
if 'confirmar_finalizacao' not in st.session_state: 
    st.session_state.confirmar_finalizacao = False
if 'nups_para_receber' not in st.session_state: 
    st.session_state.nups_para_receber = []


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
            # 1. RASTREADOR: Vai piscar amarelo na tela mostrando que o clique funcionou
            st.warning("🔄 Solicitando acesso ao banco de dados do SISAFA NAVAL...") 
            
            df_users = carregar_dados_cache(ABA_USUARIOS)
            
            if not df_users.empty:
                # --- DINÂMICA DE TAMANHO (NIP 8 | CNPJ 14) ---
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
                    senha_na_planilha = str(user_match.iloc[0, 4]).strip()
                    senha_digitada = senha.strip()

                    if senha_na_planilha == senha_digitada:
                        st.session_state.logged_in = True
                        st.session_state.user_id = u_id_limpo 
                        st.session_state.user_full_name = str(user_match.iloc[0, 1]).upper()
                        st.session_state.user_perfil = str(user_match.iloc[0, 2]).upper()
                        st.rerun()
                    else:
                        st.error("Senha incorreta.")
                        with st.expander("🔍 Detalhes do erro (Verifique sua Planilha)"):
                            st.write(f"**ID Encontrado:** {u_id_limpo}")
                            st.write(f"**Texto na Planilha:** `{senha_na_planilha}`")
                            st.write(f"**Texto Digitado:** `{senha_digitada}`")
                else:
                    st.error(f"Usuário {u_id_limpo} não cadastrado no sistema.")
                    
            else:
                # 2. ALARME CRÍTICO: Avisa se o sistema baixou uma tabela de usuários vazia!
                st.error("🚨 ERRO GRAVE: A tabela de usuários não foi carregada ou está completamente vazia. Verifique a conexão com o Google Sheets.")

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
    # ⚓ FORÇA A BARRA LATERAL A FICAR VISÍVEL
    with st.sidebar:
        
        # Logo com verificação de segurança
        if os.path.exists(caminho_logo): 
            st.image(caminho_logo)
        
        # Informações do Usuário (com proteção contra valores vazios)
        user_display = st.session_state.get('user_id', 'Não identificado')
        setor_display = st.session_state.get('modulo_ativo', 'Indefinido')
        
        st.markdown(f"""
            <div style='text-align:center; padding:10px; background-color:#f0f2f6; border-radius:10px;'>
                <b>ID:</b> {user_display}<br>
                <b>Setor:</b> {setor_display}
            </div>
        """, unsafe_allow_html=True)
        
        st.divider()

        # Botões de Ação
        if st.button("🔄 Trocar de Setor", use_container_width=True):
            st.session_state.modulo_ativo = None
            st.cache_data.clear() # Limpa o cache ao trocar para evitar lixo de outro setor
            st.rerun()
            
        if st.button("❌ Sair", use_container_width=True, type="secondary"):
            st.session_state.logged_in = False
            st.session_state.modulo_ativo = None
            st.rerun()

    # O cabeçalho que já está funcionando na sua imagem {3D5F733A-7B21-4666-8803-61A278617E8F}.png
    st.markdown(f'<div class="welcome-box">⚓ SISAFA-NAVAL: {st.session_state.modulo_ativo}</div>', unsafe_allow_html=True)
    
    # --- DESENVOLVIMENTO DOS MÓDULOS ESPECÍFICOS ---

    if st.session_state.modulo_ativo == "SECOM":
        st.header("📥 Cadastro de Faturas (SECOM)")
        aba_a = sh.worksheet(ABA_TABELA_A)
        oses = {r[0].strip(): r[1].strip() for r in aba_a.get_all_values()[1:] if r[0]}
        
        nup_in = st.text_input("NUP (Ex: 63060.000123/2026-10)", placeholder="63060.000000/0000-00")
        
        c1, c2 = st.columns(2)
        sel_cnpj = c1.selectbox("Selecione o CNPJ da OSE", [""] + sorted(list(oses.keys())))
        empresa_nome = oses.get(sel_cnpj, "")
        c2.text_input("Empresa (OSE)", value=empresa_nome, disabled=True)
        
        num_fatura = st.text_input("Número da Fatura (Alfanumérico ou S/N)")
        v_ap = st.number_input("Valor Apresentado (R$)", min_value=0.0, format="%.2f")

        # --- NOVA LÓGICA DE VALIDAÇÃO ---
        nup_padrao = r"^\d{5}\.\d{6}/\d{4}-\d{2}$" # Valida o formato 00000.000000/0000-00

        if st.button("CADASTRAR FATURA"):
            # 1. Checagem de campos vazios
            if not (nup_in and sel_cnpj and num_fatura and v_ap > 0):
                st.warning("⚠️ Preencha todos os campos obrigatórios antes de cadastrar.")
            
            # 2. Validação rigorosa do formato do NUP
            elif not re.match(nup_padrao, nup_in):
                st.error("❌ Formato de NUP incorreto! Use o padrão: 63060.000000/2026-00")
            
            # 3. Checagem de duplicidade (Instantânea usando o DF local)
            elif nup_in in df['nup'].astype(str).values:
                st.error(f"🚫 Opa! Nobre {st.session_state.user_full_name}, o NUP {nup_in} já foi cadastrado por outro (a) usuário (a)! 🚫")
            
            else:
                st.session_state.confirmar_secom = True

        # --- CAIXA DE CONFIRMAÇÃO ---
        if st.session_state.confirmar_secom:
            st.markdown("---")
            st.warning(f"**⚠️ CONFIRMAÇÃO:** Tem certeza de que os dados da fatura **{num_fatura}** estão corretos?")
            col_sim, col_nao = st.columns(2)
            
            if col_sim.button("✅ SIM, confirmar dados", use_container_width=True):
                with st.spinner("Efetuando registro..."):
                    try:
                        # 1. Conexão com a planilha de processos
                        aba_proc = sh.worksheet("SISAFA-NAVAL-processos")
                        
                        # --- 🛡️ BLINDAGEM ANTI-DUPLICIDADE EM TEMPO REAL ---
                        nups_na_planilha = aba_proc.col_values(2) 
                        if nup_in in nups_na_planilha:
                            st.error(f"🚫 Opa! Nobre {st.session_state.user_full_name}, o NUP {nup_in} já foi cadastrado por outro usuário! 🚫")
                            st.session_state.confirmar_secom = False
                            time.sleep(2)
                            st.rerun()

                        dt_hoje = datetime.now().strftime("%d/%m/%Y")
                        
                        # 2. Prepara e Grava a Linha de Processo
                        nova_linha = [
                            str(datetime.now().timestamp()), nup_in, sel_cnpj, empresa_nome, 
                            num_fatura, v_ap, 0, v_ap, datetime.now().month, datetime.now().year, 
                            1, st.session_state.user_id, dt_hoje, dt_hoje, "", "", "", ""
                        ]
                        aba_proc.append_row(nova_linha)
                        
                        # 3. Alimenta aba HISTORICO e LOGS (Apenas UMA vez aqui dentro)
                        registrar_historico(nup_in, num_fatura, "0", "1", v_ap, "Entrada via SECOM")
                        registrar_acao(nup_in, num_fatura, "CADASTRO_INICIAL", f"Fatura cadastrada por {st.session_state.user_full_name}")
                        
                        # --- 🚀 RESET DOS CAMPOS ---
                        # Aqui garantimos que o próximo lançamento venha limpo
                        st.session_state["input_fat_secom"] = ""
                        st.session_state["input_val_secom"] = 0.0
                        st.session_state.confirmar_secom = False
                        
                        st.success(f"🎉 Sucesso! Fatura {num_fatura} inserida. Obrigado, {st.session_state.user_full_name}! 🚀")
                        time.sleep(0.5) 
                        st.rerun()

                    except Exception as e:
                        st.error(f"Erro ao conectar com a planilha: {e}")
                        st.stop()

            if col_nao.button("❌ NÃO, voltar e corrigir", use_container_width=True):
                st.session_state.confirmar_secom = False
                st.rerun()

                

    elif st.session_state.modulo_ativo == "AUDITORIA":
        st.header("⚖️ Divisão de Auditoria em Saúde ⚕️")
        
        # 1. 🧼 Limpeza de espaços (Garante que não haja "mes_competencia ")
        df.columns = [str(col).strip() for col in df.columns]

        # 2. 🛡️ Tradução Preventiva (Garante que 'mês' com acento vire 'mes' sem acento)
        # Adicionei as variações mais comuns que causam erro
        df = df.rename(columns={
            'mês_competência': 'mes_competencia',
            'mês_competencia': 'mes_competencia',
            'mes_competência': 'mes_competencia',
            'Mês': 'mes_competencia',
            'MES': 'mes_competencia'
        })

        # Criação das 6 abas solicitadas
        t_fila, t_mesa, t_auditadas, t_busca, t_stats, t_rel = st.tabs([
            "📥 Fila de Espera", "🩺 Em Auditagem", "✅ Auditadas", 
            "🔍 Consultas", "📊 Produtividade", "💬 Relacionamento"
        ])

        # --- MAPEAMENTO DE MESES PARA EXIBIÇÃO ---
        mapa_meses = {
            1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR", 5: "MAI", 6: "JUN",
            7: "JUL", 8: "AGO", 9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ"
        }
    
        # 3. ⚙️ Criação da sigla (Agora com proteção contra erro de coluna ausente)
        if 'mes_competencia' in df.columns:
            df['mes_sigla'] = pd.to_numeric(df['mes_competencia'], errors='coerce').map(mapa_meses)
        else:
            # Se mesmo assim não achar, ele cria a coluna vazia para não travar o app
            df['mes_sigla'] = ""
            st.error("⚠️ Atenção: Coluna 'mes_competencia' não localizada na planilha!")

        # 1. ABA: FILA DE ESPERA
        with t_fila:
            # --- MANOBRA DE SEGURANÇA SISAFA ---
            # Primeiro verificamos se a coluna existe. Se existir, filtramos.
            if 'status' in df.columns:
                df_fila = df[df['status'] == 1].copy()
            else:
                # Se não existir, criamos um DataFrame vazio para o sistema não "capotar"
                df_fila = pd.DataFrame(columns=df.columns)
                st.warning("⚠️ Atenção: Coluna 'status' não localizada na planilha!")

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
                            try:
                                for n in nups_sel:
                                    # 1. Move fisicamente o status na planilha de processos
                                    mover_status(n, 2, auditor_nip=st.session_state.user_id)
                                    
                                    # 2. Captura cirúrgica dos dados da fatura para evitar NameError
                                    linha_fatura = df[df['nup'] == n].iloc[0]
                                    fat_n = str(linha_fatura['Numero_da_fatura'])
                                    v_apres = limpar_valor(linha_fatura['valor_apresentado'])
                                    auditor_nip = str(st.session_state.get('user_id', 'N/A'))
                                    
                                    # 3. Log rápido na memória de ações
                                    registrar_acao(n, fat_n, "RECEBIMENTO", f"Auditor {auditor_nip} recebeu o processo.")
                                    
                                    # 4. IMPLEMENTAÇÃO DA SUA FUNÇÃO NATIVA DE HISTÓRICO
                                    # Parâmetros: nup, fatura, origem (1), destino (2), valor, observação
                                    registrar_historico(
                                        nup=str(n),
                                        fatura=fat_n,
                                        origem="1",
                                        destino="2",
                                        valor=v_apres,
                                        obs=f"Processo recebido pelo Auditor NIP {auditor_nip}. Início da análise técnica da fatura."
                                    )
                                
                                # 5. Limpeza estratégica de cache para atualizar a fila e os KPIs na hora
                                st.cache_data.clear()
                                
                                st.toast("Sucesso! Processos movidos para 'Em Auditagem'.", icon="✅")
                                time.sleep(1)
                                st.rerun()
                                
                            except Exception as e:
                                st.error(f"Erro ao processar recebimento na Auditoria: {e}")
                    else:
                        st.warning("⚠️ Selecione ao menos um NUP.")

                st.divider()

                # --- FERRAMENTA DE CORREÇÃO (ERRO HUMANO) ---
                with st.expander("🛠️ CORRIGIR ERROS DE CADASTRO (NUP, Fatura, Valor ou Empresa)"):
                    st.write("Selecione um processo da fila para editar os dados originais:")
                    
                    nup_edit = st.selectbox("Escolha o NUP para corrigir:", [""] + df_fila['nup'].tolist(), key="sb_edit_fila")
                    
                    if nup_edit:
                        dados = df_fila[df_fila['nup'] == nup_edit].iloc[0]
                        
                        # Mapeamento do histórico completo (apenas leitura das empresas já existentes no SISAFA)
                        mapa_ose_cnpj = dict(zip(df['ose'].astype(str), df['cnpj'].astype(str)))
                        lista_oses = sorted([ose for ose in mapa_ose_cnpj.keys() if ose.strip() not in ["nan", "None", ""]])
                        
                        # Linha 1: Dados Financeiros e de Identificação
                        col_e1, col_e2, col_e3 = st.columns(3)
                        novo_nup = col_e1.text_input("Novo NUP:", value=str(dados['nup']))
                        nova_fat = col_e2.text_input("Nova Fatura:", value=str(dados['Numero_da_fatura']))
                        novo_val = col_e3.number_input("Novo Valor (R$):", value=float(dados.get('valor_limpo', 0.0)), format="%.2f")
                        
                        # Linha 2: Dados da Empresa (OSE e CNPJ estritamente amarrados)
                        col_e4, col_e5 = st.columns([2, 1])
                        
                        ose_atual = str(dados['ose'])
                        idx_ose = lista_oses.index(ose_atual) if ose_atual in lista_oses else 0
                        
                        # Caixa de seleção restrita às empresas já cadastradas na base
                        nova_ose = col_e4.selectbox("Empresa Cadastrada (OSE):", options=lista_oses, index=idx_ose)
                        
                        # Busca automaticamente o CNPJ vinculado à OSE selecionada acima
                        cnpj_sugerido = mapa_ose_cnpj.get(nova_ose, str(dados['cnpj']))
                        novo_cnpj = col_e5.text_input("CNPJ Correspondente:", value=cnpj_sugerido)
                        
                        if st.button("💾 SALVAR CORREÇÃO", use_container_width=True):
                            try:
                                aba_edit = sh.worksheet(ABA_PROCESSOS)
                                celula = aba_edit.find(nup_edit)
                                
                                if celula:
                                    # Atualiza os dados originais
                                    aba_edit.update_cell(celula.row, 2, novo_nup)
                                    aba_edit.update_cell(celula.row, 5, nova_fat)
                                    aba_edit.update_cell(celula.row, 6, novo_val)
                                    aba_edit.update_cell(celula.row, 8, novo_val) 
                                    
                                    # ⚠️ Colunas 3 e 4 representam CNPJ e OSE na planilha.
                                    aba_edit.update_cell(celula.row, 3, novo_cnpj) 
                                    aba_edit.update_cell(celula.row, 4, nova_ose)
                                    
                                    registrar_acao(novo_nup, nova_fat, "CORREÇÃO", f"Corrigido por {st.session_state.user_id}")
                                    
                                    st.success("✅ Dados da fatura e da empresa atualizados com sucesso no SISAFA!")
                                    st.cache_data.clear() 
                                    time.sleep(1)
                                    st.rerun()
                            except Exception as e:
                                st.error(f"Erro ao salvar: {e}")

        # 2. ABA: EM AUDITAGEM
        # =======================================================
        # 💾 BOTÃO DE RASCUNHO (SIDEBAR - AGORA FUNCIONA)
        # =======================================================
        with st.sidebar:
            st.markdown("---")
            
            if st.button("💾 SALVAR RASCUNHO", type="primary", use_container_width=True):
                # 1. Tenta pegar do "baú"
                nup_para_salvar = st.session_state.get("nup_ativo")
                
                if nup_para_salvar:
                    with st.spinner("Protegendo dados na nuvem..."):
                        # --- CAPTURA DINÂMICA DOS CENTROS DE CUSTO DIGITADOS ---
                        vals = {}
                        prefixo_input = "inp_"
                        sufixo_input = f"_{nup_para_salvar}"
                        
                        # Varre a memória ativa procurando os valores preenchidos na tela
                        for chave, valor in st.session_state.items():
                            if chave.startswith(prefixo_input) and chave.endswith(sufixo_input):
                                nome_campo = chave[len(prefixo_input):-len(sufixo_input)]
                                vals[nome_campo] = valor
                        
                        # Captura a justificativa direto da área de texto ativa
                        just = st.session_state.get("txt_just_mesa", "")
                        key_glosas = f"relatorio_glosa_{nup_para_salvar}"
                        
                        key_lista_g6_salvar = f"lista_g6_{nup_para_salvar}"
                        dados_g6_salvar = st.session_state.get(key_lista_g6_salvar, [{"tipo": "", "desc": "", "qtd": 1, "valor": 0.0}])

                        sucesso = salvar_rascunho_auditoria(
                            nup_para_salvar, 
                            st.session_state[key_glosas], 
                            vals, 
                            just,
                            dados_g6_salvar
                        )
                        if sucesso:
                            st.sidebar.success("✅ Rascunho completo salvo com sucesso! ⚓")
                        else:
                            st.sidebar.error("❌ Falha ao salvar no Google.")
                else:
                    st.sidebar.warning("⚠️ Selecione um processo na mesa de trabalho primeiro!")
        
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
            # --- MESA DE TRABALHO (VISÃO COLETIVA) ---
            st.subheader("🩺 Mesa de Trabalho da Auditoria 🐆")
            df_mesa = df[df['status'] == 2].copy()

            if df_mesa.empty:
                st.info("Não há processos em auditagem no momento.")
            else:
                st.write("**Faturas em análise técnica no setor:**")
                st.dataframe(df_mesa[['nup', 'ose', 'valor_apresentado', 'mes_sigla', 'ano_competencia', 'obs']], use_container_width=True)
                
                st.divider()
                
                nup_audit = st.selectbox("Selecione o NUP para realizar a análise:", [""] + df_mesa['nup'].tolist(), key="sb_nup_analise_mesa_final")

                # ESTA LINHA ABAIXO DEVE ESTAR ALINHADA COM O SELECTBOX ACIMA
                if nup_audit:
                    # 1. INICIALIZE AS VARIÁVEIS AQUI
                    valores_detalhados = st.session_state.get("valores_detalhados_ativo", {})
                    just_glosa = st.session_state.get("just_glosa_ativo", "")
                    
                    tabela_ref_glosa = obter_tabela_referencia_glosa()

                    dados_nup = df_mesa[df_mesa['nup'] == nup_audit].iloc[0]
                    num_fat = dados_nup['Numero_da_fatura']
                    v_apres = limpar_valor(dados_nup['valor_apresentado'])
                    
                    st.session_state["nup_ativo"] = nup_audit
                    st.session_state["valores_detalhados_ativo"] = valores_detalhados
                    st.session_state["just_glosa_ativo"] = just_glosa 
                    
                    try:
                        lista_tabela_a = carregar_dados_cache(ABA_TABELA_A)
                    except Exception as e:
                        st.error(f"Erro ao acessar a Tabela A: {e}")
                        lista_tabela_a = []


                    # --- CARREGAMENTO AUTOMÁTICO DE RASCUNHO ---
                    key_glosas = f"relatorio_glosa_{nup_audit}"
                    
                    if key_glosas not in st.session_state:
                        dados_salvos = carregar_rascunho(nup_audit)
                        if dados_salvos:
                            # 1. Recupera as Glosas
                            st.session_state[key_glosas] = dados_salvos.get('glosas', [])
                            
                            # 2. Recupera a Justificativa (Preenche a caixa de texto)
                            just_salva = dados_salvos.get('justificativa', "")
                            st.session_state["just_glosa_ativo"] = just_salva
                            st.session_state["txt_just_mesa"] = just_salva
                            
                            # 3. Recupera os Centros de Custo (Injeta os valores salvos de volta nos campos numéricos)
                            cc_salvos = dados_salvos.get('centro_custo', {})
                            st.session_state["valores_detalhados_ativo"] = cc_salvos
                            for campo, valor in cc_salvos.items():
                                # A mágica acontece aqui: associamos o valor à 'key' do number_input
                                st.session_state[f"inp_{campo}_{nup_audit}"] = float(valor)
                            
                            st.success(f"🔄 Rascunho completo recuperado! Última alteração: {nup_audit}")
                        else:
                            st.session_state[key_glosas] = [{"paciente": "", "valor": 0.0, "cod": "", "tipo": "Administrativa", "just": "", "desc_glosa": ""}]


                    st.markdown(f"#### 📝 Analisando Fatura: **{num_fat}**")
                    
                    # --- 🛠️ GAVETA DE CORREÇÃO DE NUP/VALOR ---
                    with st.expander("⚙️ Corrigir Dados Básicos (NUP ou Valor Apresentado)", expanded=False):
                        st.warning("⚠️ Altere aqui para corrigir os dados. Jamais reclame do retrabalho! Confira o Adicional de Compensação por Disponibilidade no Bilhete de Pagamento ❤️‍🔥⚓")
                        ce1, ce2 = st.columns(2)
                        
                        novo_nup = ce1.text_input("Corrigir NUP:", value=str(nup_audit), key=f"edit_nup_{nup_audit}")
                        novo_valor_apres = ce2.number_input("Corrigir Valor Apresentado (R$):", 
                                                            min_value=0.0, 
                                                                value=float(v_apres), 
                                                                format="%.2f", 
                                                                key=f"edit_v_apres_{nup_audit}")
                        
                        if st.button("💾 SALVAR CORREÇÃO NO PROCESSO", use_container_width=True):
                            with st.spinner("Atualizando base..."):
                                try:
                                    aba_proc = sh.worksheet("SISAFA-NAVAL-processos")
                                    celula = aba_proc.find(str(nup_audit))
                                    if celula:
                                        # Coluna 2 = NUP | Coluna 15 = Valor Apresentado
                                        aba_proc.update_cell(celula.row, 2, str(novo_nup))
                                        aba_proc.update_cell(celula.row, 6, str(novo_valor_apres))
                                        aba_proc.update_cell(celula.row, 8, str(novo_valor_apres))

                                        registrar_acao(nup_audit, num_fat, "CORRECAO_CADASTRO", f"NUP: {novo_nup} | Valor: {novo_valor_apres}")
                                        st.success("✅ Cadastro corrigido!")
                                        time.sleep(1)
                                        st.rerun()
                                except Exception as e:
                                    st.error(f"Erro na correção: {e}")


                    # --- DETALHAMENTO DO RELATÓRIO DE GLOSA (PADRÃO HOSBRA) ---
                    st.markdown("---")
                    st.subheader("📋 Detalhamento do Relatório de Glosa p/ paciente")
                    key_glosas = f"relatorio_glosa_{nup_audit}"

                    # Garante que a lista exista na memória com todos os campos necessários
                    if key_glosas not in st.session_state:
                        st.session_state[key_glosas] = [{"paciente": "", "valor": 0.0, "cod": "", "tipo": "Administrativa", "just": "", "desc_glosa": ""}]

                    # Puxa a tabela do cache (a função que você colocou lá no topo do arquivo)
                    tabela_ref_glosa = obter_tabela_referencia_glosa()
                    if not tabela_ref_glosa:
                        st.warning("⚠️ Tabela de Referência de Glosas indisponível. Verifique a conexão.")

                    # =======================================================
                    # TRAVA 2: FORMULÁRIO (Impede o recarregamento da tela)
                    # =======================================================
                    with st.form(key=f"form_pacientes_{nup_audit}"):
                        st.info("""
                        ⚠️ **Orientações**

                        Preencha os dados de **todos** os pacientes antes de clicar em **CONFIRMAR PACIENTES**. 
                        
                        O sistema é, de fato, lento; porém, pouco didático.  
                        
                        Contudo, não esmoreça! ⚓🇧🇷🫡
                    
                        *"Em todo trabalho há proveito; meras palavras, porém, levam à penúria." (Provérbios 14:23)*
                        
                        *"Não há substituto para o trabalho duro." (Thomas Edison)*
                        
                        *"Todo cadáver no Monte Everest já foi um dia alguém motivado, proativo e fora da sua zona de conforto" (autor desconhecido)*
                        """)
                        
                        novos_dados = []
                        
                        # Renderiza os campos de cada paciente dentro do formulário
                        for idx, item in enumerate(st.session_state[key_glosas]):
                            with st.container(border=True):
                                # 1. Ajustamos as proporções para dar espaço ao botão de excluir (col_del)
                                col_p1, col_p2, col_p3, col_del = st.columns([3, 2, 2, 1])
                                
                                temp_pac = col_p1.text_input(f"Iniciais do paciente 🤕 {idx+1}", value=item.get('paciente', ''), key=f"ini_gl_{idx}_{nup_audit}", placeholder="Ex: B.M.F.")
                                temp_val = col_p2.number_input(f"Valor R$", min_value=0.0, value=float(item.get('valor', 0.0)), key=f"v_gl_{idx}_{nup_audit}", format="%.2f")
                                
                                index_tipo = 0 if item.get('tipo', 'Administrativa') == "Administrativa" else 1
                                temp_tip = col_p3.selectbox(f"Tipo", ["Administrativa", "Técnica"], index=index_tipo, key=f"t_gl_{idx}_{nup_audit}")
                                
                                # 2. Adicionamos o Checkbox de Exclusão (Alinhado com a parte inferior)
                                col_del.markdown("<br>", unsafe_allow_html=True) # Empurra o botão para alinhar
                                temp_del = col_del.checkbox("🗑️ Excluir paciente", key=f"del_gl_{idx}_{nup_audit}")
                                
                                lista_opcoes_glosa = [""] + [f"{c} - {d}" for c, d in tabela_ref_glosa.items()]
                                
                                valor_atual_glosa = f"{item.get('cod', '')} - {item.get('desc_glosa', '')}"
                                index_glosa = 0
                                
                                if valor_atual_glosa in lista_opcoes_glosa and item.get('cod', '') != "":
                                    index_glosa = lista_opcoes_glosa.index(valor_atual_glosa)

                                escolha = st.selectbox("Código da Glosa", lista_opcoes_glosa, index=index_glosa, key=f"c_gl_{idx}_{nup_audit}")
                                temp_jus = st.text_input("Observação específica (Relatório)", value=item.get('just', ''), key=f"obs_gl_{idx}_{nup_audit}")
                                
                                temp_cod = escolha.split(" - ")[0] if escolha else ""
                                temp_desc = escolha.split(" - ")[1] if escolha else ""
                                
                                # 3. A MÁGICA DA EXCLUSÃO: Só grava na lista se NÃO marcou para excluir
                                if not temp_del:
                                    novos_dados.append({
                                        "paciente": temp_pac, 
                                        "valor": temp_val, 
                                        "cod": temp_cod, 
                                        "tipo": temp_tip, 
                                        "just": temp_jus, 
                                        "desc_glosa": temp_desc
                                    })

                            # --- NOVA LÓGICA DE BOTÕES (EM LOTE) ---
                        col_qtd, col_add, col_salvar = st.columns([1, 2, 2])
                        
                        qtd_adicionar = col_qtd.number_input("Qtd a adicionar (limite: 20 pacientes ⚠️):", min_value=1, max_value=20, value=1, step=1, key=f"qtd_add_gl_{nup_audit}")
                        btn_add = col_add.form_submit_button("➕ ADICIONAR PACIENTE NO RELATÓRIO 🤕🤧🤒")
                        btn_salvar = col_salvar.form_submit_button("💾 CONFIRMAR PACIENTES", type="primary")

                    # =======================================================
                    # AÇÕES DOS BOTÕES DO FORMULÁRIO
                    # =======================================================
                    if btn_salvar:
                        # 1. A BLINDAGEM VEM PRIMEIRO: Garante que a lista não fique vazia
                        if not novos_dados:
                            novos_dados = [{"paciente": "", "valor": 0.0, "cod": "", "tipo": "Administrativa", "just": "", "desc_glosa": ""}]
                        
                        # 2. Salva tudo de uma vez na memória principal
                        st.session_state[key_glosas] = novos_dados
                        
                        # 3. Dá o feedback e recarrega
                        st.success("✅ Pacientes confirmados e calculados com sucesso!")
                        time.sleep(1)
                        st.rerun()

                    if btn_add:
                        # Salva o que foi digitado e adiciona uma linha em branco
                        st.session_state[key_glosas] = novos_dados
                        for _ in range(qtd_adicionar):
                            st.session_state[key_glosas].append({
                            "paciente": "", "valor": 0.0, "cod": "", "tipo": "Administrativa", "just": "", "desc_glosa": ""
                            })
                        st.rerun()

                    # --- 1. RESUMO FINANCEIRO (Soma Automática dos Pacientes) ---
                    
                    # Calculamos o total somando os valores de cada paciente inserido
                    st.divider()
                    total_glosa_geral = sum(g['valor'] for g in st.session_state.get(key_glosas, []))
                    v_liquido_alvo = round(v_apres - total_glosa_geral, 2)

                    c1, c2 = st.columns(2)
                    with c1:
                        # Mostramos o total que o auditor detalhou nos pacientes
                        st.metric("Total Glosado", f"R$ {total_glosa_geral:,.2f}")
                        # Esta justificativa é a que você usará para o e-mail da OSE[cite: 2]
                        just_glosa = st.text_area("Descrição Resumida da Glosa", 
                                                height=100, 
                                                key="txt_just_mesa",
                                                help="⚠️ Essa descrição que aparecerá no corpo do e-mail e na capa da auditoria!👈")
                    with c2:
                        # O valor líquido agora é calculado com base no total detalhado
                        v_liquido_alvo = round(v_apres - total_glosa_geral, 2)
                        st.metric("Valor Apresentado", f"R$ {v_apres:,.2f}")
                        st.metric("Valor Líquido Final", f"R$ {v_liquido_alvo:,.2f}", 
                                delta=f"- R$ {total_glosa_geral:,.2f}" if total_glosa_geral > 0 else None, 
                                delta_color="inverse")


                    st.divider()
                    st.markdown("### 🔍 Detalhamento por Centro de Custo")
                    st.info(f"A soma dos campos abaixo deve ser: **R$ {v_liquido_alvo:,.2f}**")

                    # --- 2. DEFINIÇÃO DAS LISTAS (ORDEM EXATA DA PLANILHA) ---
                    g1_hosp = ["Internações UTI (exceto OPME)", "Internações não UTI (exceto OPME)", "SIAD", "HOME CARE", "Pequenas Cirurgias", "Consultas ambulatoriais", "Consultas emergenciais", "OPME", "Remédio de Alto Custo: Quimioterápicos", "Remédio de Alto Custo: Imunobiológicos", "Remédio de Alto Custo: Antibióticos"]
                    g2_lab = ["Análises Clínicas", "RX Convencional", "Tomografias", "Ressonâncias magnéticas", "Ultrassonografias"]
                    g3_spec = ["Exames oftalmológicos", "Holter 24h", "Mapa 24h", "Estudo eletrofisiológico (para estudo de arritmia cardíaca)", "Angiotomografia coronariana", "Cintilografia miocárdica", "Teste Ergométrico", "Exames do Sistema Digestório e anexos", "FACO (Catarata)", "Injeção Anti-VEGF (Ex: Lucentis)", "Revascularização miocárdica", "Angioplastia coronariana com ou sem Stent", "Cateterismo cardíaco"]
                    g4_terap = ["Hemodiálise", "Fisioterapia", "Fonoaudiologia", "Psicologia / Psicoterapia", "Avaliação neuropsicológica", "Psicopedagogia", "Terapia Ocupacional", "Musicoterapia"]
                    g5_odonto = ["Consultas", "Laboratórios Odontológicos", "Ex. Radiol. e Doc. Orto", "Prótese", "Ortodontia"]
                    
                    valores_detalhados = {}

                    def header_audit(texto, cor_fundo, cor_txt="black"):
                        st.markdown(f'<div style="background-color:{cor_fundo};padding:8px;border-radius:5px;margin:15px 0 10px 0;"><b style="color:{cor_txt}">{texto}</b></div>', unsafe_allow_html=True)

                    # =======================================================
                    # TRAVA 3: FORMULÁRIO DE CENTROS DE CUSTO (Fim da lentidão)
                    # =======================================================
                    with st.form(key=f"form_cc_{nup_audit}"):
                        st.warning("⚠️ Preencha todos os valores abaixo e clique em 'CONFIRMAR VALORES' no final da caixa.")
                        
                        # Renderização Grupos I ao V
                        header_audit("🟦 Grupo I: Assistência Médico-Hospitalar", "#ADD8E6")
                        c_a, c_b = st.columns(2)
                        for i, campo in enumerate(g1_hosp):
                            target = c_a if i % 2 == 0 else c_b
                            valores_detalhados[campo] = target.number_input(campo, min_value=0.0, format="%.2f", key=f"inp_{campo}_{nup_audit}")

                        header_audit("🟪 Grupo II: Exames laboratoriais e radiológicos", "#4B0082", "white")
                        c_a, c_b = st.columns(2)
                        for i, campo in enumerate(g2_lab):
                            target = c_a if i % 2 == 0 else c_b
                            valores_detalhados[campo] = target.number_input(campo, min_value=0.0, format="%.2f", key=f"inp_{campo}_{nup_audit}")

                        header_audit("🌸 Grupo III: Exames por especialidade", "#FFB6C1")
                        c_a, c_b = st.columns(2)
                        for i, campo in enumerate(g3_spec):
                            target = c_a if i % 2 == 0 else c_b
                            valores_detalhados[campo] = target.number_input(campo, min_value=0.0, format="%.2f", key=f"inp_{campo}_{nup_audit}")

                        header_audit("🟩 Grupo IV: Procedimentos terapêuticos", "#90EE90")
                        c_a, c_b = st.columns(2)
                        for i, campo in enumerate(g4_terap):
                            target = c_a if i % 2 == 0 else c_b
                            valores_detalhados[campo] = target.number_input(campo, min_value=0.0, format="%.2f", key=f"inp_{campo}_{nup_audit}")

                        header_audit("🟨 Grupo V: Assistência odontológica", "#FFFF00")
                        c_a, c_b = st.columns(2)
                        for i, campo in enumerate(g5_odonto):
                            target = c_a if i % 2 == 0 else c_b
                            valores_detalhados[campo] = target.number_input(campo, min_value=0.0, format="%.2f", key=f"inp_{campo}_{nup_audit}")
                        
                        # O Botão que envia os dados de uma só vez
                        btn_cc = st.form_submit_button("💾 CONFIRMAR VALORES DOS CENTROS DE CUSTO", type="primary")
                        if btn_cc:
                            st.session_state["valores_detalhados_ativo"] = valores_detalhados
                            st.success("✅ Valores registrados na memória!")

                    # =======================================================
                    # O GRUPO VI CONTINUA AQUI FORA (SEM ALTERAÇÕES)
                    # =======================================================

                    # --- 3. GRUPO VI: OUTROS (LÓGICA CAMALEÃO) ---
                    mapa_cores_outros = {
                        "Outros medicamentos": "#ADD8E6", "Outros exames": "#E6E6FA", 
                        "Outros procedimentos (SADT)": "#90EE90", "Outros procedimentos (assistência odontológica)": "#FFFF00", 
                        "Outros custos não especificados": "#FFCC99", "Outros procedimentos oftalmológicos": "#FFB6C1", 
                        "Outros procedimentos cardiológicos": "#FFB6C1", "Outros exames cardiológicos": "#FFB6C1"
                    }

                    header_audit("⬜ Grupo VI: Outros", "#D3D3D3")
                
                    key_lista_g6 = f"lista_g6_{nup_audit}"
                    if key_lista_g6 not in st.session_state:
                        st.session_state[key_lista_g6] = [{"tipo": "", "desc": "", "qtd": 1, "valor": 0.0}]

                    # =======================================================
                    # TRAVA 4: FORMULÁRIO DO GRUPO VI (Fim da sobrecarga da API)
                    # =======================================================
                    with st.form(key=f"form_g6_{nup_audit}"):
                        st.info("⚠️ Registe os custos adicionais e prima 'CONFIRMAR ITENS' para atualizar o cálculo.")
                        
                        novos_dados_g6 = []
                        
                        # Renderiza os campos de entrada para cada item
                        for idx, item in enumerate(st.session_state[key_lista_g6]):
                            with st.container(border=True):
                                c1, c2 = st.columns([4, 1])
                                
                                # Lógica para manter o tipo selecionado na memória
                                val_tipo = item.get("tipo", "")
                                idx_tipo = 0
                                lista_tipos = [""] + list(mapa_cores_outros.keys())
                                if val_tipo in lista_tipos:
                                    idx_tipo = lista_tipos.index(val_tipo)
                                
                                temp_tipo = c1.selectbox(f"Tipo de custo extra {idx+1}:", lista_tipos, index=idx_tipo, key=f"t_g6_{idx}_{nup_audit}")
                                
                                # A MÁGICA DA EXCLUSÃO DENTRO DO FORMULÁRIO (Igual aos pacientes)
                                c2.markdown("<br>", unsafe_allow_html=True)
                                temp_del = c2.checkbox("🗑️ Excluir", key=f"del_g6_{idx}_{nup_audit}")
                                
                                if temp_tipo:
                                    cor_viva = mapa_cores_outros[temp_tipo]
                                    st.markdown(f'<div style="background-color:{cor_viva};padding:5px;border-radius:5px;margin-bottom:10px;"><b style="color:black;">Lançamento em: {temp_tipo}</b></div>', unsafe_allow_html=True)
                                
                                temp_desc = st.text_input("Descrição detalhada:", value=item.get("desc", ""), key=f"desc_g6_{idx}_{nup_audit}")
                                cq1, cq2 = st.columns(2)
                                temp_qtd = cq1.number_input("Quantidade:", min_value=1, value=int(item.get("qtd", 1)), step=1, key=f"qtd_g6_{idx}_{nup_audit}")
                                temp_val = cq2.number_input("Custo Total (R$):", min_value=0.0, value=float(item.get("valor", 0.0)), format="%.2f", key=f"val_g6_{idx}_{nup_audit}")
                                
                                # Só grava na lista nova se a caixa de excluir NÃO estiver marcada
                                if not temp_del:
                                    novos_dados_g6.append({
                                        "tipo": temp_tipo, 
                                        "desc": temp_desc, 
                                        "qtd": temp_qtd, 
                                        "valor": temp_val
                                    })
                        
                        # --- NOVA LÓGICA DE BOTÕES (EM LOTE) ---
                        col_qtd_g6, col_add_g6, col_salvar_g6 = st.columns([1, 2, 2])
                        
                        qtd_add_g6 = col_qtd_g6.number_input("Qtd a adicionar (limite: 20 unidades de custo ⚠️):", min_value=1, max_value=20, value=1, step=1, key=f"qtd_add_g6_{nup_audit}")
                        btn_add_g6 = col_add_g6.form_submit_button("➕ ADICIONAR OUTRO ITEM")
                        btn_salvar_g6 = col_salvar_g6.form_submit_button("💾 CONFIRMAR ITENS (GRUPO VI)", type="primary")

                    # =======================================================
                    # AÇÕES DOS BOTÕES DO FORMULÁRIO GRUPO VI
                    # =======================================================
                    if btn_salvar_g6:
                        # Se o utilizador apagou todos os itens, mantém uma linha vazia por precaução
                        if not novos_dados_g6:
                            novos_dados_g6 = [{"tipo": "", "desc": "", "qtd": 1, "valor": 0.0}]
                            
                        st.session_state[key_lista_g6] = novos_dados_g6
                        st.success("✅ Grupo VI calculado com sucesso!")
                        time.sleep(1)
                        st.rerun()

                    if btn_add_g6:
                        # Salva o que já foi digitado e acrescenta uma nova linha vazia
                        st.session_state[key_lista_g6] = novos_dados_g6
                        for _ in range(qtd_add_g6):
                            st.session_state[key_lista_g6].append({"tipo": "", "desc": "", "qtd": 1, "valor": 0.0})
                        st.rerun()



                    # --- 4. VALIDAÇÃO MATEMÁTICA ---
                    total_g6 = sum(it["valor"] for it in st.session_state[key_lista_g6])                   
                    soma_geral = round(sum(valores_detalhados.values()) + total_g6, 2)
                    diferenca = round(v_liquido_alvo - soma_geral, 2)

                    st.divider()
                    if diferenca == 0:
                        st.success(f"✅ A soma bateu! (R$ {soma_geral:,.2f}😬🍾🎊💯👏👏👏👏👏👏)")
                        trava_cc = False
                    else:
                        st.error(f"❌ Diferença: R$ {diferenca:,.2f} (Total itens: R$ {soma_geral:,.2f})")
                        trava_cc = True

                    # Conferência de E-mail
                    with st.container(border=True):
                        try:
                            cnpj_fat = str(dados_nup['cnpj']).strip().split('.')[0]
                            df_ose = carregar_dados_cache(ABA_TABELA_A)
                            linha_o = df_ose[df_ose['CNPJ'].astype(str).str.contains(cnpj_fat)]
                            df_u = carregar_dados_cache(ABA_USUARIOS)
                            match_u = df_u[df_u['NIP'].astype(str).str.strip() == str(st.session_state.user_id).strip()]
                            
                            if not match_u.empty and not linha_o.empty:
                                email_aud = match_u['Email'].values[0] if 'Email' in match_u.columns else match_u['E-mail'].values[0]
                                email_dest = linha_o['E-mail Principal da OSE'].values[0]
                                nome_ose = linha_o['Razão Social'].values[0]
                                st.write(f"🏢 **OSE:** {nome_ose} | 📩 **E-mail:** {email_dest}")
                                trava_confirmacao = st.checkbox("Confirmo os dados acima.", key=f"chk_conf_{nup_audit}")
                            else:
                                trava_confirmacao = False
                                st.error("⚠️ Dados de e-mail não localizados.")
                        except: trava_confirmacao = False

                    # --- ÁREA DE AÇÕES FINAIS ---
                    col_fin, col_mail, col_pdf = st.columns(3)

                    # 1. PREPARAÇÃO DOS DADOS PARA OS PDFs (Fora dos botões para estarem prontos para download)
                    auditor_atual = st.session_state.get('user_full_name', 'Auditor (a)')
                    auditor_nip = str(st.session_state.get('user_id', 'N/A'))
                    
                    # Busca os dados do contrato (Tabela A) para o cabeçalho do PDF
                    cnpj_busca = str(dados_nup.get('cnpj', '')).strip().split('.')[0]
                    
                    # --- SUBSTITUA O BLOCO ANTIGO POR ESTA VERSÃO PANDAS ---
                    dados_ose_contrato = {}

                    # 1. Verifica se é DataFrame e não está vazio
                    if isinstance(lista_tabela_a, pd.DataFrame) and not lista_tabela_a.empty:
                        
                        # 2. Converte a coluna CNPJ para string e filtra usando o cnpj_busca
                        # Isso é muito mais rápido e seguro que um loop 'for'
                        mask = lista_tabela_a['CNPJ'].astype(str).str.startswith(cnpj_busca)
                        df_filtrado = lista_tabela_a[mask]
                        
                        # 3. Se encontrar algo, pega a primeira linha e transforma em dicionário
                        if not df_filtrado.empty:
                            dados_ose_contrato = df_filtrado.iloc[0].to_dict()
                        else:
                            st.warning(f"⚠️ Nenhuma OSE encontrada para o CNPJ: {cnpj_busca}")
                    else:
                        st.error("⚠️ A Tabela A está vazia ou não foi carregada corretamente.")
                    
                    # Gera o número do relatório (sequencial automático)
                    num_relatorio = obter_proximo_numero_glosa()

                    # Gera os bytes dos PDFs
                    pdf_capa_bytes = gerar_relatorio_pdf(
                        dados_nup, auditor_atual, total_glosa_geral, just_glosa, 
                        valores_detalhados, [g1_hosp, g2_lab, g3_spec, g4_terap, g5_odonto],
                        st.session_state[key_lista_g6], v_apres
                    )

                    
                    pdf_glosa_bytes = gerar_relatorio_glosa_pdf(
                        dados_nup, 
                        dados_ose_contrato, 
                        st.session_state[key_glosas], 
                        {"nome": auditor_atual, "posto": "Auditor Responsavel"}, 
                        num_relatorio,
                        just_glosa  
                    )

                    # 2. BOTÕES DE DOWNLOAD (col_pdf)
                    with col_pdf:
                        st.download_button("📄 CAPA DA AUDITORIA", data=pdf_capa_bytes, file_name=f"CAPA_{num_fat}.pdf", mime="application/pdf", use_container_width=True)
                        if total_glosa_geral > 0:
                            st.download_button("📋 RELATÓRIO DE GLOSA", data=pdf_glosa_bytes, file_name=f"GLOSA_{num_relatorio}_{auditor_atual}_{num_fat}_26.pdf", mime="application/pdf", use_container_width=True)
                    
                    # --- 3. BOTÃO FINALIZAR ---
                    if col_fin.button("✅ FINALIZAR AUDITORIA", use_container_width=True, disabled=trava_cc or not trava_confirmacao):
                        if total_glosa_geral > 0 and not just_glosa:
                            st.error("⚠️ Justificativa obrigatória para glosa.")
                        else:
                            try:
                                # --- EXTRAÇÃO E NUMERAÇÃO ---
                                num_relatorio = obter_proximo_numero_relatorio(sh) # Gera o número automático aqui
                                
                                cnpj_ose = str(dados_ose_contrato.get('CNPJ', 'N/A'))
                                nome_ose = str(dados_ose_contrato.get('Razão Social', 'N/A'))
                                num_fat = str(dados_nup.get('Numero_da_fatura', 'S/N'))
                                mes_comp = int(limpar_valor(dados_nup.get('mes_competencia', 0)))
                                ano_comp = int(limpar_valor(dados_nup.get('ano_competencia', 2026)))
                                v_apres = limpar_valor(dados_nup.get('valor_apresentado', 0.0))
                                campos_todos_grupos = g1_hosp + g2_lab + g3_spec + g4_terap + g5_odonto
                                
                                # Filtro G6
                                lista_g6_bruta = st.session_state.get(key_lista_g6, [])
                                lista_g6_limpa = [it for it in lista_g6_bruta if it.get('tipo') and limpar_valor(it.get('valor', 0)) > 0]

                                # Geração do PDF com o número novo
                                pdf_glosa_bytes = gerar_relatorio_glosa_pdf(
                                    dados_nup, dados_ose_contrato, st.session_state[key_glosas], 
                                    {"nome": auditor_atual, "posto": "Auditor Responsavel"}, 
                                    num_relatorio, just_glosa
                                )

                            except Exception as e:
                                st.error(f"Erro na preparação: {e}")
                                st.stop()

                            with st.spinner("Gravando e atualizando sistemas..."):
                                try:
                                    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                    
                                    # --- GRAVAÇÃO ANALÍTICA (SISAFA-NAVAL-Auditoria) ---
                                    aba_audit_detalhe = sh.worksheet("SISAFA-NAVAL-Auditoria")
                                    todas_as_linhas = []
                                    valores_reais = [float(valores_detalhados.get(campo, 0)) for campo in campos_todos_grupos]
                                    valores_zerados = [0.0] * len(campos_todos_grupos)

                                    if not lista_g6_limpa:
                                        linha = [agora, str(nup_audit), cnpj_ose, nome_ose, num_fat, mes_comp, ano_comp] + valores_reais + ["", "", 0, 0.0, auditor_nip]
                                        todas_as_linhas.append(linha)
                                    else:
                                        item1 = lista_g6_limpa[0]
                                        linha1 = [agora, str(nup_audit), cnpj_ose, nome_ose, num_fat, mes_comp, ano_comp] + valores_reais + [str(item1['tipo']), str(item1['desc']), int(item1['qtd']), float(item1['valor']), auditor_nip]
                                        todas_as_linhas.append(linha1)
                                        for extra in lista_g6_limpa[1:]:
                                            linha_ex = [agora, str(nup_audit), cnpj_ose, nome_ose, num_fat, mes_comp, ano_comp] + valores_zerados + [str(extra['tipo']), str(extra['desc']), int(extra['qtd']), float(extra['valor']), auditor_nip]
                                            todas_as_linhas.append(linha_ex)

                                    aba_audit_detalhe.append_rows(todas_as_linhas)

                                    # --- GRAVAÇÃO DE GLOSA (Estrutura 1-para-Muitos) ---
                                    if total_glosa_geral > 0:
                                        aba_glosa_detalhe = sh.worksheet("SISAFA-NAVAL-Auditoria-glosa")
                                        lote_glosa = []
                                        for g in st.session_state[key_glosas]:
                                            if g['paciente'] and limpar_valor(g['valor']) > 0:
                                                lote_glosa.append([
                                                    agora, str(nup_audit), cnpj_ose, nome_ose,
                                                    dados_ose_contrato.get('Termo de credenciamento', 'N/A'),
                                                    dados_ose_contrato.get('Numero_edital', 'N/A'),
                                                    dados_ose_contrato.get('Validade_edital', 'N/A'),
                                                    num_fat, mes_comp, ano_comp, v_apres,
                                                    num_relatorio, # O número automático entra aqui
                                                    str(g['paciente']), limpar_valor(g['valor']),
                                                    str(g.get('desc_glosa', g.get('desc', 'N/A'))),
                                                    str(g['cod']), str(just_glosa), str(auditor_nip)
                                                ])
                                        aba_glosa_detalhe.append_rows(lote_glosa)

                                    # --- ATUALIZAÇÃO DO STATUS (CORRIGINDO O NAMEERROR) ---
                                    aba_proc = sh.worksheet("SISAFA-NAVAL-processos") # Certifique-se que o nome é aba_proc
                                    celula = aba_proc.find(str(nup_audit))
                                    if celula:
                                        aba_proc.update_cell(celula.row, 7, float(total_glosa_geral))
                                        aba_proc.update_cell(celula.row, 8, float(v_liquido_alvo))
                                        aba_proc.update_cell(celula.row, 11, 3) # Status 3
                                        aba_proc.update_cell(celula.row, 12, str(auditor_nip))
                                        aba_proc.update_cell(celula.row, 14, agora)
                                        
                                        # Histórico
                                        aba_hist = sh.worksheet("SISAFA-NAVAL-historico")
                                        # CORREÇÃO DO NameError: Use aba_hist em vez de aba_p
                                        aba_hist.append_row([agora, nup_audit, num_fat, 2, 3, auditor_nip, v_apres, just_glosa])
                                        
                                        registrar_acao(nup_audit, num_fat, "FATURA_AUDITADA", f"Auditada por {auditor_atual}")
                                        st.cache_data.clear()
                                        st.success(f"✅ Auditagem do NUP {nup_audit} finalizada! Obrigado pela paciência, {auditor_atual}!😅")
                                        time.sleep(1)
                                        st.rerun()

                                except Exception as e:
                                    st.error(f"Erro ao salvar: {e}")

                    # Botão de E-mail (Seu código original)
                    if col_mail.button("📧 ENCAMINHAR GLOSA P/ OSE", use_container_width=True, disabled=not trava_confirmacao):
                        if disparar_email_glosa(email_dest, num_fat, glosa_input, just_glosa, nome_ose, email_aud):
                            registrar_acao(nup_audit, num_fat, "EMAIL_GLOSA_ENVIADO", f"Destino: {email_dest}")
                            st.toast("E-mail enviado!", icon="✅")


                    # --- ÁREA DE AÇÕES FINAIS ---
                    col_fin, col_mail, col_pdf = st.columns(3)


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
                    use_container_width=True, hide_index=True
                )
                
                lote_selecionado = st.multiselect(
                    "Selecionar faturas para ação:", 
                    df_auditadas['nup'].tolist(), 
                    key="ms_lote_auditadas_v2"
                )
                
                col_btn1, col_btn2 = st.columns(2)

                with col_btn1:
                    if st.button("📤 ENCAMINHAR PARA EXECUÇÃO", use_container_width=True, type="primary"):
                        if lote_selecionado:
                            with st.spinner("Registrando encaminhamento..."):
                                for n in lote_selecionado:                                    
                                    # 1. Captura dados para o registro
                                    linha_f = df_auditadas[df_auditadas['nup'] == n].iloc[0]
                                    fat_n = linha_f['Numero_da_fatura']
                                    v_liq = linha_f['v_liq_limpo'] # Valor já limpo pela sua função no topo da aba
                                    
                                    # 2. Registra na aba LOGS_ACOES (Micro)
                                    registrar_acao(n, fat_n, "ENCAMINHADO_PARA_FINANCEIRO", "Processo enviado para a Execução Financeira.")
                                    
                                    # 3. Registra na aba SISAFA-NAVAL-historico (Macro)
                                    # Status Origem: 3 | Status Destino: 3, pois só vai para o 4 quando a Execução aceita
                                    registrar_historico(n, fat_n, "3", "3", v_liq, "ENVIADO PARA O FINANCEIRO")
                            
                            st.success(f"✅ {len(lote_selecionado)} faturas encaminhadas com sucesso!")
                            time.sleep(1.2)
                            st.rerun()
                        else:
                            st.warning("Selecione ao menos um processo.")

                with col_btn2:
                    if st.button("⏪ DEVOLVER PARA AJUSTE", use_container_width=True):
                        if lote_selecionado:
                            with st.spinner("Limpando registros e devolvendo para auditagem..."):
                                try:
                                    aba_audit = sh.worksheet("SISAFA-NAVAL-Auditoria")
                                    auditor_nip = str(st.session_state.get('user_id', 'N/A'))
                                    
                                    for n in lote_selecionado:
                                        # 1. Volta o status para 2 (Em Auditagem)
                                        mover_status(n, 2)
                                        
                                        # 2. Busca e DELETA o registro analítico antigo (Evita duplicidade de custos)
                                        try:
                                            celula = aba_audit.find(str(n))
                                            if celula:
                                                aba_audit.delete_rows(celula.row)
                                        except:
                                            pass 

                                        # 3. Captura cirúrgica dos dados da fatura para evitar quebra de NameError
                                        linha_fatura = df_auditadas[df_auditadas['nup'] == n].iloc[0]
                                        fat_n = str(linha_fatura['Numero_da_fatura'])
                                        v_apres = limpar_valor(linha_fatura['valor_apresentado'])
                                        
                                        # 4. Log rápido na memória de ações
                                        registrar_acao(n, fat_n, "DEVOLUCAO_PARA_AJUSTE", f"Processo retornado ao Status 2 pelo usuário {auditor_nip}")
                                        
                                        # 5. IMPLEMENTAÇÃO DA SUA FUNÇÃO NATIVA DE HISTÓRICO
                                        # Parâmetros: nup, fatura, origem (3), destino (2), valor, observação
                                        registrar_historico(
                                            nup=str(n),
                                            fatura=fat_n,
                                            origem="3",
                                            destino="2",
                                            valor=v_apres,
                                            obs=f"Processo devolvido para reanálise/ajuste técnico na Auditoria por decisão da Execução Financeira."
                                        )
                                    
                                    # --- LINHA DE LIMPEZA DE CACHE REMOVIDA DAQUI ---
                                    
                                    st.warning(f"⏪ {len(lote_selecionado)} processos retornados para Auditagem.")
                                    time.sleep(1.2)
                                    st.rerun()
                                    
                                except Exception as e:
                                    st.error(f"Erro geral ao reverter processos e gravar o histórico: {e}")
                        else:
                            st.warning("Selecione ao menos um processo para devolver.")


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
            st.write("") # Respiro visual

            st.write("")
                
            # --- RODAPÉ COM IDENTIDADE VISUAL ---
            st.markdown("""
            <div style="text-align: center; padding: 30px; border-top: 2px solid #2e6b54; margin-top: 40px; background-color: rgba(46, 107, 84, 0.05); border-radius: 0 0 15px 15px;">
                <p style="
                    color: #2e6b54; 
                    font-weight: 900; 
                    font-size: 1.8rem; 
                    letter-spacing: 3px; 
                    line-height: 1.2;
                    text-shadow: 0 0 10px #2e6b54, 0 0 20px #2e6b54, 0 0 30px #2e6b54;
                ">
                    "RESTARÁ SEMPRE MUITO O QUE FAZER"
                </p>
                <p style="
                    color: #555; 
                    font-size: 1.1rem; 
                    font-weight: 700; 
                    margin-top: -10px;
                    text-transform: uppercase;
                    letter-spacing: 1px;
                ">
                    (SEPÚLVEDA, A.C.M)
                </p>
            </div>
            """, unsafe_allow_html=True)
            
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

            st.write("")
            st.write("")

            # =======================================================
            # === SEÇÃO 6 (MOVIDA PARA CIMA): EVOLUÇÃO AUDITADA ===
            # =======================================================
            st.divider()
            st.subheader("📈 Evolução dos Valores Auditados")
            st.write("") # Respiro visual

            # Usa o dataframe já filtrado pelos selects acima
            df_sec6 = df_p.copy()
            df_sec6['status_num'] = pd.to_numeric(df_sec6['status'], errors='coerce').fillna(0).astype(int)
            df_sec6 = df_sec6[df_sec6['status_num'] >= 3].copy() # Já auditados

            if df_sec6.empty:
                st.info("Aguardando faturas atingirem o status de Auditada para gerar este gráfico no período selecionado.")
            else:
                df_sec6['v_liq'] = df_sec6['valor_liquido'].apply(limpar_valor)
                df_sec6['v_glosa'] = df_sec6['glosa'].apply(limpar_valor)

                cats_oficiais = ["1. OSE Civis", "2. HFA", "3. Base Op. Especiais (FUSEX)", "4. BAAN", "5. HFAB"]

                def categorizar_rigido(nome):
                    n = str(nome).upper()
                    if "HOSPITAL DAS FORÇAS ARMADAS" in n or "HFA" in n: return cats_oficiais[1]
                    elif "160098" in n or "OPERAÇÕES ESPECIAIS" in n: return cats_oficiais[2]
                    elif "120624" in n or "BASE AÉREA DE ANÁPOLIS" in n: return cats_oficiais[3]
                    elif "120096" in n or "HFAB" in n: return cats_oficiais[4]
                    return cats_oficiais[0]

                df_sec6['Categoria_Audit_Final'] = df_sec6['ose'].apply(categorizar_rigido)

                df_sec6['sort_key'] = df_sec6['ano_competencia'] * 100 + df_sec6['mes_competencia']
                
                # --- Dicionário Tradutor de Meses ---
                mapa_meses_abrev = {
                    1: 'JAN', 2: 'FEV', 3: 'MAR', 4: 'ABR', 5: 'MAI', 6: 'JUN',
                    7: 'JUL', 8: 'AGO', 9: 'SET', 10: 'OUT', 11: 'NOV', 12: 'DEZ'
                }
                
                df_sec6['Competência'] = df_sec6.apply(
                    lambda x: f"{mapa_meses_abrev[int(x['mes_competencia'])]}/{str(int(x['ano_competencia']))[2:]}", axis=1
                )
                
                df_cronologico = df_sec6[['sort_key', 'Competência']].drop_duplicates().sort_values('sort_key')
                lista_competencias = df_cronologico['Competência'].tolist()

                import plotly.graph_objects as go
                fig_audit = go.Figure()

                cores_audit = {
                    "1. OSE Civis": "#1abc9c", "2. HFA": "#1e3d33", 
                    "3. Base Op. Especiais (FUSEX)": "#2e6b54", "4. BAAN": "#529471", 
                    "5. HFAB": "#76b996", "Glosa Total": "#e74c3c"
                }

                ordem_empilhamento = ["1. OSE Civis", "5. HFAB", "4. BAAN", "3. Base Op. Especiais (FUSEX)", "2. HFA"]
                
                for cat in ordem_empilhamento:
                    y_vals = [df_sec6[(df_sec6['Categoria_Audit_Final'] == cat) & (df_sec6['Competência'] == comp)]['v_liq'].sum() for comp in lista_competencias]
                    fig_audit.add_trace(go.Bar(
                        name=cat, x=lista_competencias, y=y_vals, marker_color=cores_audit[cat], 
                        offsetgroup="Liquido", hovertemplate="<b>%{x}</b><br>%{data.name}: R$ %{y:,.2f}<extra></extra>"
                    ))

                y_glosa = [df_sec6[df_sec6['Competência'] == comp]['v_glosa'].sum() for comp in lista_competencias]
                fig_audit.add_trace(go.Bar(
                    name="Glosa Total", x=lista_competencias, y=y_glosa, marker_color=cores_audit["Glosa Total"], 
                    offsetgroup="Glosa", hovertemplate="<b>%{x}</b><br>Glosa Total: R$ %{y:,.2f}<extra></extra>"
                ))

                total_liq_mensal = [df_sec6[df_sec6['Competência'] == comp]['v_liq'].sum() for comp in lista_competencias]
                fig_audit.add_trace(go.Scatter(
                    x=lista_competencias, y=total_liq_mensal, name="Tendência", mode='lines+markers',
                    line=dict(color='black', width=3, dash='dot'), marker=dict(size=8, color='black')
                ))

                fig_audit.update_layout(
                    barmode='stack',
                    hovermode="x unified", paper_bgcolor='white', plot_bgcolor='white',
                    margin=dict(t=30, b=30), # Margens ajustadas para dar fôlego
                    legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="center", x=0.5),
                    xaxis=dict(showgrid=False, linecolor='black', linewidth=1),
                    yaxis=dict(title="Montante (R$)", tickprefix="R$ ", gridcolor='rgba(0,0,0,0.05)')
                )

                st.plotly_chart(fig_audit, use_container_width=True)

            st.write("")
            st.write("")

            # =======================================================
            # === ANÁLISE DE CENTROS DE CUSTO (ORÇAMENTO) ===
            # =======================================================
            st.divider()
            st.subheader("🏢 Análise por Centro de Custo")
            st.write("") 

            try:
                aba_auditoria = sh.worksheet("SISAFA-NAVAL-Auditoria")
                df_aud = pd.DataFrame(aba_auditoria.get_all_records())
                
                if df_aud.empty:
                    st.info("Aguardando inserção de dados na aba de Auditoria para carregar os Centros de Custo.")
                else:
                    df_aud.columns = df_aud.columns.astype(str).str.strip()
                    df_aud.columns = [" ".join(col.split()) for col in df_aud.columns]
                    
                    if 'nup' in df_aud.columns:
                        df_aud['nup'] = df_aud['nup'].astype(str).str.strip()
                    
                    df_aud['sort_key'] = df_aud['ano_competencia'] * 100 + df_aud['mes_competencia']
                    
                    df_aud['Competência'] = df_aud.apply(
                        lambda x: f"{mapa_meses_abrev[int(x['mes_competencia'])]}/{str(int(x['ano_competencia']))[2:]}", axis=1
                    )
                    
                    for col in df_aud.columns:
                        if 'custo total' in col.lower():
                            df_aud.rename(columns={col: 'Outros'}, inplace=True)
                    
                    lista_oficial = [
                        "Internações UTI (exceto OPME)", "Internações não UTI (exceto OPME)", "SIAD", "HOME CARE",
                        "Pequenas Cirurgias", "Consultas ambulatoriais", "Consultas emergenciais", "OPME",
                        "Remédio de Alto Custo: Quimioterápicos", "Remédio de Alto Custo: Imunobiológicos",
                        "Remédio de Alto Custo: Antibióticos", "Análises Clínicas", "RX Convencional", "Tomografias",
                        "Ressonâncias magnéticas", "Ultrassonografias", "Exames oftalmológicos", "Holter 24h", "Mapa 24h",
                        "Estudo eletrofisiológico (para estudo de arritmia cardíaca)", "Angiotomografia coronariana",
                        "Cintilografia miocárdica", "Teste Ergométrico", "Exames do Sistema Digestório e anexos",
                        "FACO (Catarata)", "Injeção Anti-VEGF (Ex: Lucentis)", "Revascularização miocárdica",
                        "Angioplastia coronariana com ou sem Stent", "Cateterismo cardíaco", "Hemodiálise", "Fisioterapia",
                        "Fonoaudiologia", "Psicologia / Psicoterapia", "Avaliação neuropsicológica", "Psicopedagogia",
                        "Terapia Ocupacional", "Musicoterapia", "Consultas", "Laboratórios Odontológicos",
                        "Ex. Radiol. e Doc. Orto", "Prótese", "Ortodontia", "Outros"
                    ]
                    
                    colunas_centros_custo = [col for col in lista_oficial if col in df_aud.columns]

                    def limpar_custo_auditoria(val):
                        if pd.isna(val) or val == '': return 0.0
                        if isinstance(val, (int, float)): return float(val)
                        v_str = str(val).replace('R$', '').strip()
                        if not v_str: return 0.0
                        
                        if '.' in v_str and ',' in v_str:
                            if v_str.rfind(',') > v_str.rfind('.'): 
                                v_str = v_str.replace('.', '').replace(',', '.')
                            else: 
                                v_str = v_str.replace(',', '')
                        elif ',' in v_str:
                            v_str = v_str.replace(',', '.') 
                        
                        try:
                            return float(v_str)
                        except:
                            return 0.0

                    for col in colunas_centros_custo:
                        df_aud[col] = df_aud[col].apply(limpar_custo_auditoria)

                    df_long = pd.melt(
                        df_aud, id_vars=['Competência', 'sort_key', 'ose', 'nup'], 
                        value_vars=colunas_centros_custo, var_name='Centro de Custo', value_name='Valor'
                    )
                    
                    df_long = df_long[df_long['Valor'] > 0].copy()

                    if df_long.empty:
                        st.warning("Nenhum lançamento financeiro maior que R$ 0,00 encontrado nos centros de custo oficiais.")
                    else:
                        st.markdown("#### 📊 Distribuição Global de Gastos")
                        
                        df_pizza = df_long.groupby('Centro de Custo')['Valor'].sum().reset_index()
                        df_pizza = df_pizza.sort_values('Valor', ascending=False)
                        
                        fig_pie = px.pie(
                            df_pizza, names='Centro de Custo', values='Valor', hole=0.45, 
                            color_discrete_sequence=px.colors.qualitative.Prism 
                        )
                        fig_pie.update_traces(
                            textposition='inside', textinfo='percent',
                            hovertemplate="<b>%{label}</b><br>Gasto: R$ %{value:,.2f}<br>Representação: %{percent}<extra></extra>"
                        )
                        fig_pie.update_layout(
                            paper_bgcolor='white', plot_bgcolor='white',
                            margin=dict(l=20, r=20, t=30, b=20), legend=dict(font=dict(color='black'))
                        )
                        st.plotly_chart(fig_pie, use_container_width=True)

                        st.write("")
                        st.write("")

                        st.markdown("#### 📈 Evolução Individual por Centro de Custo")
                        
                        lista_centros_ativos = sorted(df_long['Centro de Custo'].unique().tolist())
                        centro_selecionado = st.selectbox(
                            "🎯 Selecione a linha de serviço para visualização técnica:", lista_centros_ativos
                        )

                        df_evol_individual = df_long[df_long['Centro de Custo'] == centro_selecionado]
                        df_evol_individual = df_evol_individual.groupby(['Competência', 'sort_key'])['Valor'].sum().reset_index()
                        df_evol_individual = df_evol_individual.sort_values('sort_key')

                        cor_tatica = '#2c5d71' 

                        fig_area_ind = px.area(
                            df_evol_individual, x='Competência', y='Valor', 
                            title=f"Histórico de Desembolso: {centro_selecionado}", markers=True, template="plotly_white"
                        )
                        fig_area_ind.update_traces(
                            line_color=cor_tatica, fillcolor='rgba(44, 93, 113, 0.2)', 
                            marker=dict(size=8, color=cor_tatica, line=dict(width=2, color="white")),
                            hovertemplate="<b>%{x}</b><br>Gasto: R$ %{y:,.2f}<extra></extra>"
                        )
                        fig_area_ind.update_layout(
                            hovermode="x unified", paper_bgcolor='white', plot_bgcolor='white', margin=dict(l=20, r=20, t=50, b=20),
                            xaxis=dict(showgrid=False, type='category', linecolor='black', tickfont=dict(color='black')),
                            yaxis=dict(title="Total Gasto (R$)", tickprefix="R$ ", gridcolor='rgba(0,0,0,0.05)', tickfont=dict(color='black'))
                        )
                        st.plotly_chart(fig_area_ind, use_container_width=True)

                        with st.expander("📋 Ver Matriz Contábil Completa de Centros de Custo por Mês"):
                            df_matrix = df_long.pivot_table(
                                index='Centro de Custo', columns='Competência', values='Valor', aggfunc='sum'
                            ).fillna(0)
                            df_matrix['Total Acumulado'] = df_matrix.sum(axis=1)
                            df_matrix = df_matrix.sort_values('Total Acumulado', ascending=False)
                            st.dataframe(df_matrix.style.format("R$ {:,.2f}"), use_container_width=True)

            # --- CORREÇÃO AQUI: O fechamento do try dos centros de custo ---
            except Exception as e:
                st.error(f"Erro ao processar a inteligência de centros de custo: {e}")

            st.write("")
            st.write("")

            
            # =======================================================
            # === NOVA SEÇÃO: INTELIGÊNCIA E MAPEAMENTO DE GLOSAS ===
            # =======================================================
            st.divider()
            st.subheader("✂️ Análise Estratégica de Glosas")
            st.write("") 

            try:
                # Carregando as duas abas
                aba_glosa = sh.worksheet("SISAFA-NAVAL-Auditoria-glosa")
                aba_glosa_ref = sh.worksheet("SISAFA-NAVAL-Tabela-de-referencia-de-glosa")
                
                df_glosa = pd.DataFrame(aba_glosa.get_all_records())
                df_glosa_ref = pd.DataFrame(aba_glosa_ref.get_all_records())

                if df_glosa.empty:
                    st.info("Aguardando inserção de registros de glosa para gerar as métricas de retenção.")
                else:
                    # Padronização de colunas
                    df_glosa.columns = df_glosa.columns.astype(str).str.strip()
                    df_glosa_ref.columns = df_glosa_ref.columns.astype(str).str.strip()

                    # Aplicação da sua função de limpeza consolidada
                    if 'Valor_glosa' in df_glosa.columns:
                        df_glosa['Valor_glosa_num'] = df_glosa['Valor_glosa'].apply(limpar_valor)
                    else:
                        df_glosa['Valor_glosa_num'] = 0.0

                    # Filtros Globais (Obedece a seleção de Ano e Mês do cabeçalho)
                    if 'ano_competencia' in df_glosa.columns and ano_sel != "Todos":
                        df_glosa = df_glosa[pd.to_numeric(df_glosa['ano_competencia'], errors='coerce') == ano_sel]
                    
                    if 'mes_competencia' in df_glosa.columns and mes_sel != "Todos":
                        mapa_mes_inverso = {1: 'JAN', 2: 'FEV', 3: 'MAR', 4: 'ABR', 5: 'MAI', 6: 'JUN', 7: 'JUL', 8: 'AGO', 9: 'SET', 10: 'OUT', 11: 'NOV', 12: 'DEZ'}
                        df_glosa['mes_sigla_glosa'] = pd.to_numeric(df_glosa['mes_competencia'], errors='coerce').map(mapa_mes_inverso)
                        df_glosa = df_glosa[df_glosa['mes_sigla_glosa'] == mes_sel]

                    if df_glosa.empty or df_glosa['Valor_glosa_num'].sum() == 0:
                        st.warning("Nenhum valor de glosa encontrado para os filtros selecionados.")
                    else:
                        # Mapeamento do Código para a Descrição Bonita
                        df_glosa['Cod_glosa'] = df_glosa['Cod_glosa'].astype(str).str.strip()
                        df_glosa_ref['Cod_glosa'] = df_glosa_ref['Cod_glosa'].astype(str).str.strip()
                        
                        mapa_desc_glosa = dict(zip(df_glosa_ref['Cod_glosa'], df_glosa_ref['Desc_glosa']))
                        
                        df_glosa['Motivo_Glosa'] = df_glosa['Cod_glosa'].astype(str) + " - " + df_glosa['Cod_glosa'].map(mapa_desc_glosa).fillna("Descrição não encontrada")

                        # --- GRÁFICO 1: Glosas por Motivo (Pizza Bonitão) ---
                        st.markdown("#### 🚫 Representatividade Financeira por Motivo de Glosa")
                        
                        df_motivos = df_glosa.groupby('Motivo_Glosa')['Valor_glosa_num'].sum().reset_index()
                        df_motivos = df_motivos[df_motivos['Valor_glosa_num'] > 0].sort_values('Valor_glosa_num', ascending=False)
                        
                        fig_glosa_motivo = px.pie(
                            df_motivos, names='Motivo_Glosa', values='Valor_glosa_num', hole=0.45,
                            color_discrete_sequence=px.colors.qualitative.Pastel
                        )
                        fig_glosa_motivo.update_traces(
                            textposition='inside', textinfo='percent',
                            hovertemplate="<b>%{label}</b><br>Valor Retido: R$ %{value:,.2f}<br>Proporção: %{percent}<extra></extra>"
                        )
                        fig_glosa_motivo.update_layout(
                            paper_bgcolor='white', plot_bgcolor='white', margin=dict(l=20, r=20, t=30, b=20),
                            legend=dict(font=dict(color='black', size=11), orientation="h", yanchor="top", y=-0.1)
                        )
                        st.plotly_chart(fig_glosa_motivo, use_container_width=True)

                        # --- GRÁFICO 2: Top 10 OSEs que mais sofrem Glosa ---
                        st.markdown("#### 🏥 Top 10 OSEs com Maior Retenção Financeira")
                        
                        df_ose_glosa = df_glosa.groupby('ose')['Valor_glosa_num'].sum().reset_index()
                        df_ose_glosa = df_ose_glosa[df_ose_glosa['Valor_glosa_num'] > 0].sort_values('Valor_glosa_num', ascending=True).tail(10)
                        
                        fig_glosa_ose = px.bar(
                            df_ose_glosa, x='Valor_glosa_num', y='ose', orientation='h', text='Valor_glosa_num',
                            color_discrete_sequence=['#c0392b']
                        )
                        fig_glosa_ose.update_traces(
                            texttemplate='R$ %{text:,.2f}', textposition='outside',
                            hovertemplate="<b>%{y}</b><br>Glosado: R$ %{x:,.2f}<extra></extra>"
                        )
                        fig_glosa_ose.update_layout(
                            hovermode="y unified", paper_bgcolor='white', plot_bgcolor='white', margin=dict(l=20, r=20, t=30, b=20),
                            xaxis=dict(title="Montante Glosado (R$)", showgrid=True, gridcolor='rgba(0,0,0,0.05)', tickprefix="R$ "),
                            yaxis_title=None
                        )
                        fig_glosa_ose.update_xaxes(range=[0, df_ose_glosa['Valor_glosa_num'].max() * 1.3]) 
                        
                        st.plotly_chart(fig_glosa_ose, use_container_width=True)

            except Exception as e:
                st.error(f"Erro ao processar os dados de Glosa: {e}")

            # =======================================================
            # --- 1. SEÇÃO: PIZZAS E VOLUMES (UM EMBAIXO DO OUTRO) ---
            # =======================================================
            st.divider()
            st.subheader("📌 Visão Geral do Volume")
            st.write("") # Respiro visual
            
            # --- PLACAR DE COMANDO: MÉTRICAS VERTICAIS ---
            with st.container():
                st.metric("Processos Cadastrados", f"{len(df_p):,}")
                st.write("") # Espaço entre os blocos
                
            with st.container():
                st.metric("Volume Total Apresentado", f"R$ {df_p['v_ap_num'].sum():,.2f}")
                st.write("")
                
            with st.container():
                st.metric("Volume Total Glosado (Retido)", f"R$ {df_p['glosa_num'].sum():,.2f}")

            st.write("") # Espaço generoso entre blocos
            st.write("")
            
            # --- PIZZA 1: SITUAÇÃO DE MESA ---
            st_counts = df_p[df_p['status'].isin([2, 3])]['status'].map({2:'Em Mesa', 3:'Concluídas'}).value_counts().reset_index()
            if not st_counts.empty:
                fig_p1 = px.pie(
                    st_counts, values='count', names='status', 
                    title="<b>Produtividade:</b> Faturas em Mesa de Auditoria vs. Concluídas", 
                    hole=0.45,
                    color_discrete_sequence=['#3498db', '#2ecc71'] # Azul técnico e Verde sucesso
                )
                fig_p1.update_layout(
                    paper_bgcolor='white', plot_bgcolor='white',
                    title_x=0.1, # Alinha o título de forma elegante
                    margin=dict(t=40, b=40),
                    legend=dict(font=dict(color='black', size=12))
                )
                st.plotly_chart(fig_p1, use_container_width=True)
            
            st.write("") # Respiro antes da próxima informação
            st.write("")

            # --- PIZZA 2: IMPACTO FINANCEIRO ---
            v_t = df_p['v_ap_num'].sum()
            v_g = df_p['glosa_num'].sum()
            if v_t > 0:
                fig_p2 = px.pie(
                    values=[v_t - v_g, v_g], names=['Líquido Aprovado', 'Glosa (Corte Técnico)'], 
                    title="<b>Análise de Impacto:</b> Proporção Financeira Retida por Glosa", 
                    hole=0.45, 
                    color_discrete_sequence=['#2e6b54', '#d32f2f'] # Verde tático e Vermelho alerta
                )
                fig_p2.update_traces(textposition='inside', textinfo='percent')
                fig_p2.update_layout(
                    paper_bgcolor='white', plot_bgcolor='white',
                    title_x=0.1,
                    margin=dict(t=40, b=40),
                    legend=dict(font=dict(color='black', size=12))
                )
                st.plotly_chart(fig_p2, use_container_width=True)

            st.write("")
            st.write("")

            # --- PIZZA 3: TOP 10 OSEs ---
            top_ose = df_p.groupby('ose')['v_ap_num'].sum().sort_values(ascending=False).head(10).reset_index()
            if not top_ose.empty:
                fig_p3 = px.pie(
                    top_ose, values='v_ap_num', names='ose', 
                    title="<b>Concentração de Risco:</b> Top 10 OSEs por Maior Volume Financeiro Apresentado",
                    hole=0.45,
                    color_discrete_sequence=px.colors.qualitative.Prism # Paleta moderna e bem definida
                )
                fig_p3.update_traces(textposition='inside', textinfo='percent')
                fig_p3.update_layout(
                    paper_bgcolor='white', plot_bgcolor='white',
                    title_x=0.1,
                    margin=dict(t=40, b=40),
                    legend=dict(font=dict(color='black', size=11)) # Tamanho ideal para não cortar os nomes das OSEs
                )
                st.plotly_chart(fig_p3, use_container_width=True)

            st.write("")
            st.write("")

            # =======================================================
            # --- 2. SEÇÃO: TERMÔMETRO E SAÚDE ---
            # =======================================================
            st.divider()
            st.subheader("🌡️ Termômetro de Saúde do Processo (Global)")
            st.write("") # Respiro visual
            
            hoje = datetime.now()
            df_p['dt_ent'] = pd.to_datetime(df_p['data_entrada'], dayfirst=True, errors='coerce')
            df_p['dias_hoje'] = (hoje - df_p['dt_ent']).dt.days

            # FUNÇÃO CORRIGIDA: Blinda contra dados vazios
            def classificar_global(d):
                if pd.isna(d): return None # Ignora processos sem data de entrada cadastrada
                if d <= 15: return "🟢 Aceitável"
                if d <= 25: return "🟡 Atenção"
                return "🔴 Em Atraso"

            df_p['situacao'] = df_p['dias_hoje'].apply(classificar_global)
            
            # Remove os vazios antes de contar
            df_saude_limpo = df_p.dropna(subset=['situacao'])
            saude_counts = df_saude_limpo['situacao'].value_counts().reset_index()
            
            if not saude_counts.empty:
                # Ordena o eixo X para fazer sentido tático: Aceitável -> Atenção -> Atraso
                ordem_saude = ["🟢 Aceitável", "🟡 Atenção", "🔴 Em Atraso"]
                
                fig_saude = px.bar(saude_counts, x='situacao', y='count', color='situacao', 
                                   title="Saúde do Passivo (Tempo desde o Cadastro na SECOM)",
                                   color_discrete_map={"🟢 Aceitável": "#2e6b54", "🟡 Atenção": "#f1c40f", "🔴 Em Atraso": "#e74c3c"},
                                   category_orders={"situacao": ordem_saude}) # Trava a ordem visual
                
                fig_saude.update_layout(margin=dict(t=40, b=20))
                st.plotly_chart(fig_saude, use_container_width=True)

            
            # =======================================================
            # === APOIO À DIRETORIA DE SAÚDE DA MARINHA (DSM) ===
            # =======================================================
            st.divider()

            # 🎨 INJEÇÃO DO CSS NEON AQUI (Garante o visual)
            st.markdown("""
            <style>
            .neon-card { background: #ffffff; border-radius: 12px; padding: 20px; margin-bottom: 20px; text-align: center; box-shadow: 0 4px 15px rgba(0, 0, 0, 0.05); transition: transform 0.2s; border-bottom: 4px solid #eee; }
            .neon-card:hover { transform: translateY(-3px); }
            .nc-title { font-size: 0.9rem; font-weight: 700; color: #555; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 10px; }
            .nc-value { font-size: 2.2rem; font-weight: 900; margin-bottom: 5px; }
            .nc-sub { font-size: 0.85rem; font-weight: 600; color: #777; }
            .card-cyan { border-bottom-color: #00e5ff; box-shadow: 0 10px 20px rgba(0, 229, 255, 0.15); }
            .card-purple { border-bottom-color: #d500f9; box-shadow: 0 10px 20px rgba(213, 0, 249, 0.15); }
            .card-alert { border-bottom-color: #ff1744; box-shadow: 0 10px 20px rgba(255, 23, 68, 0.15); background: #fffcfc;}
            .card-green { border-bottom-color: #00e676; box-shadow: 0 10px 20px rgba(0, 230, 118, 0.15); }
            .card-orange { border-bottom-color: #ff9100; box-shadow: 0 10px 20px rgba(255, 145, 0, 0.15); }
            </style>
            """, unsafe_allow_html=True)

            col_dsm1, col_dsm2 = st.columns([1, 4])
            with col_dsm1:
                try:
                    if 'caminho_escudo_dsm' in locals() and os.path.exists(caminho_escudo_dsm):
                        st.image(caminho_escudo_dsm, width=150) # Ajustado para 100px para não estourar a tela
                    else:
                        st.warning("Escudo DSM não encontrado.")
                except:
                    st.write("⚓")
            with col_dsm2:
                st.subheader("Apoio às informações prestadas à Diretoria de Saúde da Marinha (DSM)")

            # ==========================================
            # 🧹 MOTOR DE NORMALIZAÇÃO E LIMPEZA PROFUNDA (TÁTICA INDUSTRIAL RESTAURADA)
            # ==========================================
            # 1. Limpeza brutal do cabeçalho
            df_aud.columns = [str(c).strip() for c in df_aud.columns]
            
            # 2. Resgate do Método Industrial (Índices 51 = Qtd, 52 = Custo Outros)
            idx_qtd = 51
            idx_val = 52
            
            # Pega o nome exato das colunas usando a posição na planilha
            col_qtd = df_aud.columns[idx_qtd] if len(df_aud.columns) > idx_qtd else None
            col_custo_outros = df_aud.columns[idx_val] if len(df_aud.columns) > idx_val else None

            # 3. Limpeza das Strings vitais
            if 'Grupo' in df_aud.columns: df_aud['Grupo'] = df_aud['Grupo'].astype(str).str.strip()
            if 'Descrição' in df_aud.columns: df_aud['Descrição'] = df_aud['Descrição'].astype(str).str.strip()
            
            if col_qtd and col_qtd in df_aud.columns:
                df_aud[col_qtd] = pd.to_numeric(df_aud[col_qtd], errors='coerce').fillna(0).astype(int)

            # 4. Motor Blindado de Datas
            def parse_timestamp_blindado(val):
                v_str = str(val).strip()
                if not v_str or v_str.lower() in ['nan', 'none', 'nat']: return pd.NaT
                try:
                    if '-' in v_str and v_str.find('-') == 4: return pd.to_datetime(v_str, errors='coerce')
                    return pd.to_datetime(v_str, dayfirst=True, errors='coerce')
                except:
                    return pd.to_datetime(v_str, errors='coerce')

            df_aud['Data_Auditoria'] = df_aud['timestamp'].apply(parse_timestamp_blindado)
            
            mapa_meses = {1: 'JAN', 2: 'FEV', 3: 'MAR', 4: 'ABR', 5: 'MAI', 6: 'JUN', 7: 'JUL', 8: 'AGO', 9: 'SET', 10: 'OUT', 11: 'NOV', 12: 'DEZ'}
            df_aud['Mês_Auditoria'] = df_aud['Data_Auditoria'].apply(lambda d: f"{mapa_meses[d.month]}/{d.year}" if pd.notna(d) else "Desconhecido")
            df_aud['mes_competencia'] = pd.to_numeric(df_aud['mes_competencia'], errors='coerce')
            df_aud['ano_competencia'] = pd.to_numeric(df_aud['ano_competencia'], errors='coerce')
            df_aud['Competência'] = df_aud.apply(lambda r: f"{mapa_meses[int(r['mes_competencia'])]}/{str(int(r['ano_competencia']))[2:]}" if pd.notna(r['mes_competencia']) else "S/C", axis=1)
            df_aud['sort_comp'] = df_aud['ano_competencia'].fillna(0) * 100 + df_aud['mes_competencia'].fillna(0) 

            # 5. APLICAÇÃO GLOBAL DO LIMPAR_VALOR
            colunas_oficiais = [c for c in lista_oficial if c in df_aud.columns and c != "Outros"]
            
            for c in colunas_oficiais:
                df_aud[c] = df_aud[c].apply(limpar_valor)
                
            if col_custo_outros and col_custo_outros in df_aud.columns:
                df_aud[col_custo_outros] = df_aud[col_custo_outros].apply(limpar_valor)

            # 6. CRIAÇÃO DA COLUNA MESTRA (Soma Tudo: Oficiais + Outros)
            soma_oficiais = df_aud[colunas_oficiais].sum(axis=1)
            soma_outros = df_aud[col_custo_outros].fillna(0) if col_custo_outros else 0
            df_aud['Valor_Total_Auditado'] = soma_oficiais + soma_outros

            # ==========================================
            # 🎛️ FILTROS E PREPARAÇÃO
            # ==========================================
            df_datas_validas = df_aud.dropna(subset=['Data_Auditoria']).sort_values('Data_Auditoria')
            meses_audit_unicos = df_datas_validas['Mês_Auditoria'].unique().tolist()
            if "Desconhecido" in df_aud['Mês_Auditoria'].unique(): meses_audit_unicos.append("Desconhecido")

            periodo_sel = st.multiselect("🎯 Selecione o Mês da Realização da Auditoria (Filtro DSM):", meses_audit_unicos, default=meses_audit_unicos[-1:] if meses_audit_unicos else [])

            if not periodo_sel:
                st.warning("Selecione pelo menos um mês de auditoria para visualizar os dados DSM.")
            else:
                df_dsm = df_aud[df_aud['Mês_Auditoria'].isin(periodo_sel)].copy()
                
                # --- MATRIZ OFICIAL DE GRUPOS E CORES ---
                g1_hosp = ["Internações UTI (exceto OPME)", "Internações não UTI (exceto OPME)", "SIAD", "HOME CARE", "Pequenas Cirurgias", "Consultas ambulatoriais", "Consultas emergenciais", "OPME", "Remédio de Alto Custo: Quimioterápicos", "Remédio de Alto Custo: Imunobiológicos", "Remédio de Alto Custo: Antibióticos"]
                g2_lab = ["Análises Clínicas", "RX Convencional", "Tomografias", "Ressonâncias magnéticas", "Ultrassonografias"]
                g3_spec = ["Exames oftalmológicos", "Holter 24h", "Mapa 24h", "Estudo eletrofisiológico (para estudo de arritmia cardíaca)", "Angiotomografia coronariana", "Cintilografia miocárdica", "Teste Ergométrico", "Exames do Sistema Digestório e anexos", "FACO (Catarata)", "Injeção Anti-VEGF (Ex: Lucentis)", "Revascularização miocárdica", "Angioplastia coronariana com ou sem Stent", "Cateterismo cardíaco"]
                g4_terap = ["Hemodiálise", "Fisioterapia", "Fonoaudiologia", "Psicologia / Psicoterapia", "Avaliação neuropsicológica", "Psicopedagogia", "Terapia Ocupacional", "Musicoterapia"]
                g5_odonto = ["Consultas", "Laboratórios Odontológicos", "Ex. Radiol. e Doc. Orto", "Prótese", "Ortodontia"]

                mapa_cc_classe = {}
                for c in g1_hosp: mapa_cc_classe[c] = "card-cyan"
                for c in g2_lab: mapa_cc_classe[c] = "card-purple"
                for c in g3_spec: mapa_cc_classe[c] = "card-orange"
                for c in g4_terap: mapa_cc_classe[c] = "card-green"
                for c in g5_odonto: mapa_cc_classe[c] = "card-alert"

                mapa_cores_outros = {
                    "Outros medicamentos": "#ADD8E6", "Outros exames": "#E6E6FA", 
                    "Outros procedimentos (SADT)": "#90EE90", "Outros procedimentos (assistência odontológica)": "#FFFF00", 
                    "Outros custos não especificados": "#FFCC99", "Outros procedimentos oftalmológicos": "#FFB6C1", 
                    "Outros procedimentos cardiológicos": "#FFB6C1", "Outros exames cardiológicos": "#FFB6C1"
                }

                def obter_cor_outros(nome_grupo):
                    nome_limpo = str(nome_grupo).lower().strip()
                    for k, v in mapa_cores_outros.items():
                        if k.lower() in nome_limpo: return v
                    return "#ff9100" 

                # =======================================================
                # 1. RENDERIZAÇÃO DOS 42 CENTROS DE CUSTO OFICIAIS
                # =======================================================
                st.markdown("### 📋 Distribuição por Centros de Custos ao estilo DSM ⚕️")
                
                centros_ativos = [c for c in colunas_oficiais if df_dsm[c].sum() > 0]

                if not centros_ativos:
                    st.info("Nenhum valor auditado nos Centros de Custo oficiais no mês selecionado.")
                else:
                    for i in range(0, len(centros_ativos), 3):
                        cols = st.columns(3)
                        for j, centro in enumerate(centros_ativos[i:i+3]):
                            df_centro = df_dsm[df_dsm[centro] > 0]
                            valor_cc = df_centro[centro].sum()
                            
                            brk_df = df_centro.groupby(['sort_comp', 'Competência'])[centro].sum().reset_index()
                            brk_df = brk_df[brk_df[centro] > 0].sort_values('sort_comp')
                            breakdown_html = "".join([f"• <span style='color:#555;'>{row['Competência']}:</span> <b>R$ {row[centro]:,.2f}</b><br>" for _, row in brk_df.iterrows()])
                            
                            css_class = mapa_cc_classe.get(centro, "card-cyan")

                            with cols[j]:
                                st.markdown(f"""
                                <div class="neon-card {css_class}">
                                    <div class="nc-title" style="min-height: 40px;">{centro}</div>
                                    <div class="nc-value" style="font-size: 1.5rem;">R$ {valor_cc:,.2f}</div>
                                    <hr style="margin: 8px 0; border: 0.5px solid #eee;">
                                    <div class="nc-sub" style="text-align: left; font-size: 0.8rem;">
                                        <span style="font-weight:bold; color:#444;">Discriminação do Centro (Por Competência):</span><br>
                                        {breakdown_html}
                                    </div>
                                </div>
                                """, unsafe_allow_html=True)

                st.write("")
                
                # =======================================================
                # 2. RENDERIZAÇÃO DOS GRUPOS DE "OUTROS"
                # =======================================================
                st.markdown("### 📂 Grupos de Custos Diversos (O famigerado grupo Outros 😱🫨)")

                # Filtra as linhas que de fato pertencem a "Outro"
                df_outros = df_dsm[df_dsm['Grupo'].str.contains("Outro", case=False, na=False)].copy()

                # Removemos a trava do nome da coluna. Confiamos apenas no conteúdo do df_outros.
                if df_outros.empty:
                    st.info("Nenhum lançamento no Grupo VI para o período selecionado.")
                else:
                    def categorizar_grupo(texto):
                        t = str(texto).lower()
                        if "medicamento" in t: return "Outros medicamentos"
                        if "exame" in t and "cardi" not in t and "oftal" not in t: return "Outros exames"
                        if "sadt" in t or ("procedimento" in t and "odont" not in t and "oftal" not in t and "card" not in t): return "Outros procedimentos (SADT)"
                        if "odontol" in t: return "Outros procedimentos (assistência odontológica)"
                        if "oftal" in t: return "Outros procedimentos oftalmológicos"
                        if "cardi" in t and "procedimento" in t: return "Outros procedimentos cardiológicos"
                        if "cardi" in t and "exame" in t: return "Outros exames cardiológicos"
                        return "Outros custos não especificados"

                    df_outros['Grupo_Consolidado'] = df_outros['Grupo'].apply(categorizar_grupo)
                    grupos_unicos = sorted(df_outros['Grupo_Consolidado'].unique().tolist())
                    
                    for i in range(0, len(grupos_unicos), 3):
                        cols_outros = st.columns(3)
                        for j, g_nome in enumerate(grupos_unicos[i:i+3]):
                            df_g = df_outros[df_outros['Grupo_Consolidado'] == g_nome]
                            cor_hex = obter_cor_outros(g_nome)
                            
                            # Soma utilizando a coluna identificada pelo índice 52
                            total_grupo = df_g[col_custo_outros].sum()
                            
                            # Agregação por Descrição
                            if col_qtd and col_qtd in df_g.columns:
                                agg_desc = df_g.groupby('Descrição').agg({col_qtd: 'sum', col_custo_outros: 'sum'}).reset_index()
                                desc_html = "".join([f"• {row['Descrição']} ({int(row[col_qtd])} un): R$ {row[col_custo_outros]:,.2f}<br>" for _, row in agg_desc.iterrows()])
                            else:
                                agg_desc = df_g.groupby('Descrição').agg({col_custo_outros: 'sum'}).reset_index()
                                desc_html = "".join([f"• {row['Descrição']}: R$ {row[col_custo_outros]:,.2f}<br>" for _, row in agg_desc.iterrows()])
                            
                            brk = df_g.groupby('Competência')[col_custo_outros].sum()
                            breakdown_html = "".join([f"• <span style='color:#555;'>{k}:</span> <b>R$ {v:,.2f}</b><br>" for k, v in brk.items()])

                            with cols_outros[j]:
                                st.markdown(f"""
                                <div class="neon-card" style="border-bottom-color: {cor_hex}; box-shadow: 0 10px 20px {cor_hex}33;">
                                    <div class="nc-title" style="color: {cor_hex}; min-height: 40px;">{g_nome}</div>
                                    <div class="nc-value" style="font-size: 1.5rem;">R$ {total_grupo:,.2f}</div>
                                    <hr style="margin: 8px 0; border: 0.5px solid #eee;">
                                    <div class="nc-sub" style="text-align: left; font-size: 0.75rem; max-height: 200px; overflow-y: auto;">
                                        <span style="font-weight:bold; color:#444;">Itens Compilados:</span><br>
                                        {desc_html}
                                        <br><span style="font-weight:bold; color:#444;">Por Competência:</span><br>
                                        {breakdown_html}
                                    </div>
                                </div>
                                """, unsafe_allow_html=True)

                # =======================================================
                # 3. NOVA SEÇÃO: VALORES AUDITADOS POR EMPRESA (OSE)
                # =======================================================
                st.write("") 
                st.divider()
                st.markdown("### 🏢 Análise de Valores Auditados por Empresa (OSE)")
                
                if 'ose' not in df_dsm.columns:
                    st.warning("Coluna de empresa ('ose') não localizada.")
                else:
                    df_empresas = df_dsm.copy()
                    df_empresas['ose'] = df_empresas['ose'].astype(str).str.strip()
                    df_empresas = df_empresas[~df_empresas['ose'].isin(['nan', 'None', '', 'NaT'])]
                    
                    # Filtra apenas empresas cujo total (Coluna Mestra) seja maior que zero
                    df_empresas = df_empresas[df_empresas['Valor_Total_Auditado'] > 0]
                    
                    if df_empresas.empty:
                        st.info("Nenhum valor auditado por empresa no período selecionado.")
                    else:
                        empresas_unicas = sorted(df_empresas['ose'].unique().tolist())
                        
                        for i in range(0, len(empresas_unicas), 3):
                            cols_emp = st.columns(3)
                            for j, empresa in enumerate(empresas_unicas[i:i+3]):
                                df_e = df_empresas[df_empresas['ose'] == empresa]
                                
                                # Usamos a Coluna Mestra consolidada que jamais falha
                                total_empresa = df_e['Valor_Total_Auditado'].sum()
                                
                                brk_emp = df_e.groupby(['sort_comp', 'Competência'])['Valor_Total_Auditado'].sum().reset_index()
                                brk_emp = brk_emp.sort_values('sort_comp')
                                
                                breakdown_emp_html = "".join([f"• <span style='color:#555;'>{row['Competência']}:</span> <b>R$ {row['Valor_Total_Auditado']:,.2f}</b><br>" 
                                                              for _, row in brk_emp.iterrows()])
                                
                                cor_hex_emp = "#1e3d59" 
                                
                                with cols_emp[j]:
                                    st.markdown(f"""
                                    <div class="neon-card" style="border-bottom-color: {cor_hex_emp}; box-shadow: 0 10px 20px {cor_hex_emp}33;">
                                        <div class="nc-title" style="color: {cor_hex_emp}; min-height: 40px; font-size: 0.85rem;">{empresa}</div>
                                        <div class="nc-value" style="font-size: 1.5rem;">R$ {total_empresa:,.2f}</div>
                                        <hr style="margin: 8px 0; border: 0.5px solid #eee;">
                                        <div class="nc-sub" style="text-align: left; font-size: 0.75rem; max-height: 200px; overflow-y: auto;">
                                            <span style="font-weight:bold; color:#444;">Discriminação por Competência:</span><br>
                                            {breakdown_emp_html}
                                        </div>
                                    </div>
                                    """, unsafe_allow_html=True)




    # --- 6. ABA: RELACIONAMENTO (Módulo Auditoria) ---
        with t_rel:
            st.subheader("💬 Central de Relacionamento")
            
            try:
                # 1. Carregamos as mensagens
                aba_msg = sh.worksheet(ABA_MENSAGENS)
                dados_brutos = aba_msg.get_all_records()
                df_msg = pd.DataFrame(dados_brutos)
                
                if df_msg.empty:
                    st.info("Nenhuma mensagem registrada no sistema.")
                else:
                    # --- FILTRO POR SETOR ---
                    # Mostra apenas o que é destinado à Auditoria
                    df_msg_auditoria = df_msg[df_msg['setor_destino'] == "AUDITORIA"].copy()

                    if df_msg_auditoria.empty:
                        st.info("✅ Nenhuma mensagem pendente para a Auditoria.")
                    else:
                        # 2. Métricas do Setor
                        pendentes = len(df_msg_auditoria[df_msg_auditoria['status_msg'] == 'PENDENTE'])
                        
                        c1, c2 = st.columns(2)
                        c1.metric("Mensagens Recebidas", len(df_msg_auditoria))
                        c2.metric("📩 Aguardando Resposta", pendentes, delta_color="inverse")
                        
                        st.divider()

                        # 3. Inbox de Mensagens (Usando os nomes das suas 10 colunas)
                        st.write("**📥 Mensagens Destinadas à Auditoria:**")
                        
                        # Ordenamos para mostrar as PENDENTES no topo
                        df_exibir = df_msg_auditoria.sort_values(by='status_msg', ascending=False)
                        
                        st.dataframe(
                            df_exibir[['Numero_da_fatura', 'nup', 'remetente', 'data_envio', 'status_msg']], 
                            use_container_width=True,
                            hide_index=True
                        )

                        st.markdown("---")

                        # 4. Área de Resposta Técnica
                        st.markdown("### ✍️ Responder à OSE")
                        
                        # Criamos uma lista de seleção: "Fatura 123 (NUP: 000...)"
                        df_msg_auditoria['label_selecao'] = (
                            "Fatura: " + df_msg_auditoria['Numero_da_fatura'].astype(str) + 
                            " | ID: " + df_msg_auditoria['id_mensagem'].astype(str)
                        )
                        
                        selecao = st.selectbox("Selecione a mensagem para responder:", [""] + df_msg_auditoria['label_selecao'].tolist())
                        
                        if selecao:
                            # Localizamos os dados da mensagem selecionada
                            msg_data = df_msg_auditoria[df_msg_auditoria['label_selecao'] == selecao].iloc[0]
                            id_msg_alvo = str(msg_data['id_mensagem'])
                            
                            with st.container(border=True):
                                st.write(f"🏢 **Remetente (OSE):** {msg_data['remetente']}")
                                st.write(f"📑 **NUP vinculado:** {msg_data['nup']}")
                                st.info(f"💬 **Mensagem da OSE:**\n\n{msg_data['texto']}")
                                
                                resposta_texto = st.text_area("Resposta Oficial do Auditor:", height=150, placeholder="Digite o parecer técnico...")
                                
                                if st.button("📤 ENVIAR RESPOSTA PARA A OSE", use_container_width=True):
                                    if resposta_texto:
                                        with st.spinner("Gravando resposta na planilha..."):
                                            # --- AÇÃO 1: Registrar no Log Geral (Tabela-B) ---
                                            registrar_acao(msg_data['nup'], msg_data['Numero_da_fatura'], "RESPOSTA_AUDITORIA", f"Auditor respondeu ID {id_msg_alvo}")

                                            # --- AÇÃO 2: Atualizar a Aba de Mensagens ---
                                            # Localizamos a linha pelo id_mensagem (Coluna A)
                                            try:
                                                celula = aba_msg.find(id_msg_alvo)
                                                linha_idx = celula.row
                                                
                                                # Atualizamos: data_resposta (8), status_msg (9), respondido_por_nip (10)
                                                agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                                aba_msg.update_cell(linha_idx, 8, agora)
                                                aba_msg.update_cell(linha_idx, 9, "RESPONDIDO")
                                                aba_msg.update_cell(linha_idx, 10, str(st.session_state.user_id)) # NIP do Auditor logado

                                                st.success("Resposta enviada com sucesso!")
                                                time.sleep(1.5)
                                                st.rerun()
                                            except Exception as err:
                                                st.error(f"Erro ao localizar linha para resposta: {err}")
                                    else:
                                        st.warning("Por favor, escreva a resposta antes de enviar.")

            except Exception as e:
                st.error(f"Erro ao carregar Central de Relacionamento: {e}")    


    elif "EXECUÇÃO" in st.session_state.modulo_ativo:
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
                    df_fila_aud[['nup','ose','Numero_da_fatura','valor_liquido','mes_sigla','ano_competencia','dias_espera']],
                    use_container_width=True,
                    key="df_aud_recepcao"
                )
                
                nups_aud_sel = st.multiselect("Selecionar faturas auditadas para receber:", df_fila_aud['nup'].tolist(), key="ms_aud_recep")
                
                # --- COLUNAS DE AÇÃO ---
                # --- COLUNAS DE AÇÃO (3 Colunas agora) ---
                c_rec, c_cor, c_dev = st.columns(3)

                with c_rec:
                    if st.button("✅ Receber Faturas", key="btn_aud_recep", use_container_width=True):
                        if nups_aud_sel:
                            with st.spinner("Recebendo e registrando no histórico..."):
                                try:
                                    # O laço varre as faturas selecionadas no lote
                                    for n in nups_aud_sel:
                                        # 1. Evolui o processo para Status 4 (Aguard. NE)
                                        mover_status(n, 4) 
                                        
                                        # 2. Captura cirúrgica dos dados da fatura para evitar NameError
                                        linha_fatura = df[df['nup'] == n].iloc[0]
                                        fat_n = str(linha_fatura['Numero_da_fatura'])
                                        v_apres = limpar_valor(linha_fatura['valor_apresentado'])
                                        
                                        # 3. Log rápido na memória de ações
                                        registrar_acao(n, fat_n, "RECEBIMENTO_FINANCEIRO", "Fatura recebida pela Execução Financeira.")
                                        
                                        # 4. CHAMADA DA SUA FUNÇÃO NATIVA DE HISTÓRICO (Perfeita!)
                                        # Parâmetros: nup, fatura, origem (3), destino (4), valor, observação
                                        registrar_historico(
                                            nup=str(n), 
                                            fatura=fat_n, 
                                            origem="3", 
                                            destino="4", 
                                            valor=v_apres, 
                                            obs="Fatura aceita e recebida pelo setor de Execução Financeira."
                                        )
                                    
                                    # 5. Limpa o cache para atualizar o Painel Tático no mesmo instante
                                    st.cache_data.clear()
                                    
                                    st.success(f"✅ {len(nups_aud_sel)} faturas recebidas com sucesso!")
                                    time.sleep(1)
                                    st.rerun()
                                    
                                except Exception as e:
                                    st.error(f"Erro geral no recebimento das faturas: {e}")
                        else:
                            st.warning("Selecione faturas.")

                with c_cor:
                    if st.button("⚠️ Corrigir Processo", key="btn_aud_corr", use_container_width=True):
                        if len(nups_aud_sel) == 1:
                            st.session_state['modo_correcao'] = nups_aud_sel[0]
                        elif len(nups_aud_sel) > 1:
                            st.error("Corrija um por vez.")
                        else:
                            st.warning("Selecione um processo.")

                with c_dev:
                    if st.button("⏪ Devolver p/ Auditoria", key="btn_aud_dev", use_container_width=True, type="secondary"):
                        if nups_aud_sel:
                            with st.spinner("Devolvendo..."):
                                try:
                                    # 1. Abre a conexão com a aba de histórico fora do loop (economiza cota)
                                    aba_hist = sh.worksheet("SISAFA-NAVAL-historico")
                                    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                    auditor_nip = str(st.session_state.get('user_id', 'N/A'))
                                    
                                    for n in nups_aud_sel:
                                        # Move fisicamente o processo de volta para "Em Auditagem"
                                        mover_status(n, 2) 
                                        
                                        # Captura cirúrgica dos dados reais da fatura no DataFrame
                                        linha_fatura = df[df['nup'] == n].iloc[0]
                                        fat_n = str(linha_fatura['Numero_da_fatura'])
                                        v_apres = limpar_valor(linha_fatura['valor_apresentado'])
                                        
                                        # Log Rápido
                                        registrar_acao(n, fat_n, "DEVOLUCAO_AUDITORIA", "Fatura devolvida pela Execução para reanálise.")
                                        
                                        # --- RESTAURAÇÃO DO HISTÓRICO MILITAR ---
                                        # Registra a retroação do Status 3 (Fila Execução) -> Status 2 (Em Auditagem)
                                        aba_hist.append_row([
                                            agora, 
                                            str(n), 
                                            fat_n, 
                                            3, # Status Antigo
                                            2, # Status Novo
                                            auditor_nip, 
                                            v_apres, 
                                            "Fatura devolvida pela Execução para reanálise técnica na Auditoria."
                                        ])
                                    
                                    # Limpa o cache para atualizar as estatísticas na hora
                                    st.cache_data.clear()
                                    
                                    st.warning(f"⏪ {len(nups_aud_sel)} faturas devolvidas para a Auditoria!")
                                    time.sleep(1)
                                    st.rerun()
                                    
                                except Exception as e:
                                    st.error(f"Erro ao registrar devolução no histórico: {e}")
                        else:
                            st.warning("Selecione faturas para devolver.")

                # --- FORMULÁRIO DE CORREÇÃO 
                # --- FORMULÁRIO DE CORREÇÃO  ---
                if 'modo_correcao' in st.session_state and st.session_state['modo_correcao']:
                    nup_alvo = st.session_state['modo_correcao']
                    dados_originais = df[df['nup'] == nup_alvo].iloc[0]
                    
                    with st.expander(f"🛠️ Editando Processo: {nup_alvo}", expanded=True):
                        with st.form("form_correcao_auditoria"):
                            
                            # 1. PUXA A TABELA A DO CACHE (Economia de cota garantida)
                            try:
                                df_tabela_a = carregar_dados_cache(ABA_TABELA_A)
                            except Exception as e:
                                st.error(f"Erro ao carregar dados de referência das OSEs: {e}")
                                df_tabela_a = pd.DataFrame()

                            # Prepara a lista para o Selectbox: "CNPJ - Razão Social"
                            if not df_tabela_a.empty:
                                # Cria uma coluna combinada para o usuário escolher visualmente com facilidade
                                df_tabela_a['combo_ose'] = df_tabela_a['CNPJ'].astype(str) + " - " + df_tabela_a['Razão Social'].astype(str)
                                lista_combos = df_tabela_a['combo_ose'].tolist()
                                
                                # Tenta descobrir qual é o índice do CNPJ atual do processo para já vir selecionado
                                cnpj_atual = str(dados_originais['cnpj']).strip()
                                index_padrao = 0
                                for idx, item in enumerate(lista_combos):
                                    if item.startswith(cnpj_atual):
                                        index_padrao = idx
                                        break
                            else:
                                lista_combos = [""]
                                index_padrao = 0

                            c1, c2 = st.columns(2)
                            novo_nup = c1.text_input("Corrigir NUP:", value=str(dados_originais['nup']))
                            
                            # 2. O SELECTBOX INTELIGENTE (Substitui o text_input antigo)
                            escolha_ose = c2.selectbox(
                                "Selecionar Novo CNPJ/OSE:", 
                                options=lista_combos, 
                                index=index_padrao,
                                help="Selecione a OSE correta. O sistema mapeará o CNPJ (Col 3) e o Nome (Col 4) automaticamente."
                            )
                            
                            # 3. EXTRAÇÃO DOS DADOS DA SELEÇÃO (Sem bater no Google Sheets)
                            if escolha_ose and " - " in escolha_ose:
                                novo_cnpj = escolha_ose.split(" - ")[0].strip()
                                novo_nome_ose = escolha_ose.split(" - ")[1].strip()
                            else:
                                novo_cnpj = str(dados_originais['cnpj'])
                                novo_nome_ose = str(dados_originais['ose'])

                            # Exibe um campo informativo desabilitado para o usuário ver o nome da OSE mudando em tempo real
                            st.text_input("Nome da OSE Vinculada (Coluna 4):", value=novo_nome_ose, disabled=True)

                            st.divider()
                            
                            c1, c2 = st.columns(2)
                            nova_fat = c1.text_input("Corrigir Nº Fatura:", value=str(dados_originais['Numero_da_fatura']))
                            v_apres_edit = c2.number_input("Valor Apresentado (R$):", value=float(limpar_valor(dados_originais['valor_apresentado'])))
                            v_liquido_edit = c1.number_input("Valor Líquido (R$):", value=float(limpar_valor(dados_originais['valor_liquido'])))
                            
                            nova_obs = st.text_area("Justificativa da alteração:", placeholder="Descreva o que foi corrigido.")

                            btn_save, btn_cancel = st.columns(2)
                            
                            if btn_save.form_submit_button("💾 Aplicar Correções", use_container_width=True):
                                try:
                                    with st.spinner("Atualizando base de dados..."):
                                        aba_proc = sh.worksheet("SISAFA-NAVAL-processos")
                                        row = aba_proc.find(str(nup_alvo)).row
                                        
                                        # Atualização em lote ou célula a célula (Mantendo a segurança de colunas)
                                        aba_proc.update_cell(row, 2, str(novo_nup))       # Coluna B (2): NUP
                                        aba_proc.update_cell(row, 3, str(novo_cnpj))      # Coluna C (3): CNPJ
                                        aba_proc.update_cell(row, 4, str(novo_nome_ose))  # Coluna D (4): Nome da OSE (Mapeamento automático realizado!)
                                        aba_proc.update_cell(row, 5, str(nova_fat))       # Coluna E (5): Numero_da_fatura
                                        aba_proc.update_cell(row, 6, f"R${v_apres_edit:,.2f}") # Coluna F (6): Valor Apresentado
                                        aba_proc.update_cell(row, 8, f"R${v_liquido_edit:,.2f}") # Coluna H (8): Valor Líquido
                                        
                                        registrar_acao(novo_nup, nova_fat, "CORRECAO_DADOS", f"CNPJ alterado para {novo_cnpj} ({novo_nome_ose}) | Justificativa: {nova_obs}")
                                        
                                        # Limpa o cache do dataframe de processos para a alteração refletir na tela imediatamente
                                        st.cache_data.clear()
                                        
                                        st.success("✅ Processo e OSE vinculada atualizados com sucesso!")
                                        st.session_state['modo_correcao'] = None
                                        time.sleep(1)
                                        st.rerun()
                                except Exception as e:
                                    st.error(f"Erro na atualização da planilha: {e}")
                            
                            if btn_cancel.form_submit_button("❌ Sair"):
                                st.session_state['modo_correcao'] = None
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
                # Seleção por NOTA FISCAL (Blindado contra tipos mistos)
                # 1. Transforma tudo em texto limpo e remove vazios (NaN)
                df_fila_fiscal['nf'] = df_fila_fiscal['nf'].fillna("").astype(str).str.strip()
                
                # 2. Pega os valores únicos, ignorando linhas onde não há NF, e ordena
                nfs_unicas = [nf for nf in df_fila_fiscal['nf'].unique() if nf != ""]
                lista_nfs = sorted(nfs_unicas)
                
                nfs_sel = st.multiselect("Selecione a(s) Nota(s) Fiscal(is) para aceitar:", options=lista_nfs, key="ms_nf_recep")
                
                if st.button("🚀 Aceitar e Liquidar Notas Fiscais", key="btn_nf_recep", use_container_width=True):
                    if nfs_sel:
                        # Buscamos todos os NUPs vinculados a essas NFs
                        nups_da_nf = df_fila_fiscal[df_fila_fiscal['nf'].isin(nfs_sel)]['nup'].tolist()
                        
                        with st.spinner(f"Processando {len(nups_da_nf)} faturas..."):
                            try:
                                for n in nups_da_nf:
                                    # 1. Evolui de 6 para 7 (Em Liquidação)
                                    mover_status(n, 7) 
                                    
                                    # 2. Capturas cirúrgicas no DataFrame para evitar NameError
                                    linha_fatura = df[df['nup'] == n].iloc[0]
                                    fat_n = str(linha_fatura['Numero_da_fatura'])
                                    v_apres = limpar_valor(linha_fatura['valor_apresentado'])
                                    
                                    # Tenta pescar o número da NF correspondente a este processo específico
                                    df_match_nf = df_fila_fiscal[df_fila_fiscal['nup'] == n]
                                    nf_n = str(df_match_nf['nf'].values[0]) if not df_match_nf.empty else "S/N"
                                    
                                    # 3. Registro rápido na memória de ações
                                    registrar_acao(n, fat_n, "NF_ACEITA_FINANCEIRO", f"NF {nf_n} conferida e aceita pela Execução.")
                                    
                                    # 4. IMPLEMENTAÇÃO DA SUA FUNÇÃO NATIVA DE HISTÓRICO
                                    # Parâmetros: nup, fatura, origem (6), destino (7), valor, observação
                                    registrar_historico(
                                        nup=str(n),
                                        fatura=fat_n,
                                        origem="6",
                                        destino="7",
                                        valor=v_apres,
                                        obs=f"Nota Fiscal nº {nf_n} aceita e homologada pelo financeiro. Processo enviado para liquidação no SIAFI."
                                    )
                                
                                # 5. Limpeza estratégica de cache para acender o painel tático na hora
                                st.cache_data.clear()
                                
                                st.success(f"✅ {len(nfs_sel)} Notas Fiscais aceitas! Processos movidos para liquidação.")
                                time.sleep(1.5)
                                st.rerun()
                                
                            except Exception as e:
                                st.error(f"Erro ao processar e registrar o histórico das NFs: {e}")
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
                            try:
                                # --- AQUI ESTÁ A CORREÇÃO: DEFININDO A ABA ---
                                aba_p = sh.worksheet("SISAFA-NAVAL-processos") 
                                aba_hist = sh.worksheet("SISAFA-NAVAL-historico")
                                agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                                for nup in nups_sel:
                                    cell = aba_p.find(str(nup)) # str() garante a busca correta
                                    if cell:
                                        # Coluna 15 é onde você guarda a NE na sua planilha
                                        aba_p.update_cell(cell.row, 15, ne_completa)
                                        
                                        # Move para Status 5 (Empenhado)
                                        mover_status(nup, 5)
                                        
                                        # --- CAPTURA CIRÚRGICA DOS DADOS DA FATURA NO DF ---
                                        linha_fatura = df[df['nup'] == nup].iloc[0]
                                        fatura_n = str(linha_fatura['Numero_da_fatura'])
                                        v_liq_num = limpar_valor(linha_fatura['valor_liquido'])
                                        auditor_nip = str(st.session_state.get('user_id', 'N/A'))
                                        
                                        # Logs e Registros Oficiais
                                        registrar_acao(nup, fatura_n, "NE_CADASTRADA", f"NE {ne_completa} vinculada ao CNPJ {cnpj_alvo}")
                                        
                                        # --- SEU REGISTRAR HISTÓRICO COM AS VARIÁVEIS VACINADAS ---
                                        # Ajustado para usar a estrutura correta sem dar NameError
                                        aba_hist.append_row([
                                            agora, 
                                            str(nup), 
                                            fatura_n, 
                                            4, 
                                            5, 
                                            auditor_nip, 
                                            v_liq_num, 
                                            f"NE {ne_completa} cadastrada para o CNPJ {cnpj_alvo}"
                                        ])

                                st.success(f"✅ Sucesso! NE {ne_completa} cadastrada para {len(nups_sel)} faturas.")
                                time.sleep(1)
                                st.rerun()
                                
                            except Exception as e:
                                st.error(f"Erro ao acessar a planilha de processos: {e}")
                            
                            # ---------------------------------------
                            
                            time.sleep(1.5)
                            st.rerun()

                st.divider()
                st.subheader("📋 Processos Disponíveis para Empenho")
                cols_v = ['nup','cnpj','ose','mes_sigla','ano_competencia','valor_liquido']
                st.dataframe(f_status_4[cols_v].rename(columns={'mes_sigla':'Mês'}), use_container_width=True)
            else:
                st.info("Não há faturas aguardando emissão de NE.")

            st.markdown("---") # Separador visual dentro da aba

            # --- SEÇÃO 2: ENVIO PARA FISCALIZAÇÃO (DENTRO DA ABA CORRETA) ---
            st.markdown("---")
            st.markdown("### 📤 2. Encaminhar para Fiscalização (ou Cancelar Empenho)")

            f_status_5 = df[df['status'] == 5].copy()

            if not f_status_5.empty:
                # 1. BUSCA DE FISCAIS
                try:
                    # Carrega e força ser um DataFrame
                    dados_usuarios = carregar_dados_cache("SISAFA-NAVAL-Usuarios")
                    df_u = pd.DataFrame(dados_usuarios)
                    
                    if not df_u.empty:
                        # --- LIMPEZA DE COLUNAS ---
                        # Tira espaços e põe tudo em MAIÚSCULO para não errar o nome
                        df_u.columns = [str(c).strip().upper() for c in df_u.columns]
                        
                        # --- VERIFICAÇÃO DINÂMICA ---
                        # Procuramos as colunas que contenham "PERFIL" e "NOME" no título
                        col_perfil = next((c for c in df_u.columns if "PERFIL" in c), None)
                        col_nome = next((c for c in df_u.columns if "NOME" in c), None)

                        if col_perfil and col_nome:
                            # Filtro: Busca os termos 'fiscalização' ou 'fiscal_global'
                            fiscais_disp = df_u[
                                df_u[col_perfil].astype(str).str.contains(
                                    "fiscalização|fiscal_global", 
                                    case=False, 
                                    na=False
                                )
                            ]
                            lista_fiscais = sorted(fiscais_disp[col_nome].unique().tolist())
                        else:
                            # Se ele não achou as colunas pelos nomes, ele avisa
                            st.error(f"⚠️ Colunas 'NOME' ou 'PERFIL' não encontradas!")
                            st.info(f"Colunas lidas na planilha: {list(df_u.columns)}")
                            lista_fiscais = []
                    else:
                        st.warning("⚠️ A planilha de usuários está vazia.")
                        lista_fiscais = []
                        
                except Exception as e:
                    st.error(f"❌ Erro ao processar lista: {e}")
                    lista_fiscais = []
               
                # 2. INTERFACE DE SELEÇÃO
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
                        key="sb_fiscal_destino"
                    )

                # --- BOTÕES DE AÇÃO ---
                col_env, col_canc = st.columns(2)

                # AÇÃO A: ENCAMINHAR (STATUS 5 -> 6)
                if col_env.button("📧 Encaminhar p/ Fiscalização", use_container_width=True, type="primary"):
                    if selecao_ne and fiscal_destinatario:
                        nups_para_enviar = f_status_5[f_status_5['ne'].isin(selecao_ne)]['nup'].tolist()
                        with st.spinner("Encaminhando..."):
                            for nup in nups_para_enviar:
                                mover_status(nup, 6)
                                dados_n = f_status_5[f_status_5['nup'] == nup].iloc[0]
                                registrar_acao(nup, dados_n['Numero_da_fatura'], "ENVIO_FISCALIZACAO", f"Fiscal: {fiscal_destinatario}")
                                registrar_historico(nup, dados_n['Numero_da_fatura'], "5", "6", dados_n['valor_apresentado'], f"Enviado p/ {fiscal_destinatario}")
                            
                            st.success(f"✅ Sucesso! Empenhos enviados para {fiscal_destinatario}.")
                            time.sleep(1.2)
                            st.rerun()
                    else:
                        st.warning("⚠️ Selecione a NE e o Fiscal de destino.")

                # AÇÃO B: CANCELAR EMPENHO (STATUS 5 -> 4)
                if col_canc.button("🚫 Cancelar Nota de Empenho", use_container_width=True):
                    if selecao_ne:
                        nups_para_cancelar = f_status_5[f_status_5['ne'].isin(selecao_ne)]['nup'].tolist()
                        with st.spinner("Cancelando e retornando status..."):
                            # Acessamos a aba de processos pelo nome exato informado
                            aba_proc = sh.worksheet("SISAFA-NAVAL-processos") 
                            
                            for nup in nups_para_cancelar:
                                dados_n = f_status_5[f_status_5['nup'] == nup].iloc[0]
                                ne_velha = dados_n['ne']
                                fat_n = dados_n['Numero_da_fatura']
                                v_momento = dados_n['valor_apresentado']
                                
                                # 1. Retorna para Status 4 (Aguardando NE)
                                mover_status(nup, 4) 
                                
                                # 2. Localiza e apaga a NE (Coluna 15)
                                try:
                                    celula_nup = aba_proc.find(str(nup))
                                    if celula_nup:
                                        aba_proc.update_cell(celula_nup.row, 15, "") # Coluna 15 é a 'ne'
                                except Exception as e:
                                    st.error(f"Erro ao limpar NE na planilha: {e}")

                                # 3. Atualiza Logs e Histórico conforme as colunas solicitadas
                                registrar_acao(nup, fat_n, "CANCELAMENTO_NE", f"NE {ne_velha} cancelada. Retorno ao Status 4.")
                                registrar_historico(nup, fat_n, "5", "4", v_momento, f"Cancelamento de NE {ne_velha}")

                            st.error(f"🚫 {len(nups_para_cancelar)} faturas retornaram para o Status 4.")
                            time.sleep(1.2)
                            st.rerun()
                    else:
                        st.warning("⚠️ Selecione as NEs para cancelar.")

                # --- TABELA DE VISUALIZAÇÃO ---
                st.subheader("📊 Faturas Empenhadas aguardando envio")
                mapa_meses = {1:"JAN", 2:"FEV", 3:"MAR", 4:"ABR", 5:"MAI", 6:"JUN", 7:"JUL", 8:"AGO", 9:"SET", 10:"OUT", 11:"NOV", 12:"DEZ"}
                f_status_5['mes_sigla'] = f_status_5['mes_competencia'].map(mapa_meses)
                
                cols_f = ['ne', 'ose', 'nup', 'valor_liquido', 'mes_sigla', 'ano_competencia']
                st.dataframe(f_status_5[cols_f].sort_values(by='ne'), use_container_width=True, hide_index=True)

            else:
                st.info("Não há Notas de Empenho aguardando envio para fiscalização.")



        # --- ABA 3: GESTÃO DE PAGAMENTOS (Liquidação e Pagamento) ---
        with tab3:
            mapa_meses = {1:"JAN", 2:"FEV", 3:"MAR", 4:"ABR", 5:"MAI", 6:"JUN", 
                          7:"JUL", 8:"AGO", 9:"SET", 10:"OUT", 11:"NOV", 12:"DEZ"}

            # =================================================================
            # SEÇÃO 1: LIQUIDAÇÃO (Status 7 -> 8)
            # =================================================================
            st.markdown("### 🛠️ 1. Liquidar Faturas (por NF)")
            st.write("Selecione as Notas Fiscais que foram liquidadas.")

            # Filtrar faturas no status 7 e que possuem NF preenchida
            df_status_7 = df[df['status'] == 7].copy()
            
            if not df_status_7.empty:
                df_status_7['mes_sigla'] = df_status_7['mes_competencia'].map(mapa_meses)
                # Garante que não apareçam NFs vazias no seletor
                lista_nfs_7 = df_status_7['nf'].dropna().unique().tolist()

                nfs_para_liquidar = st.multiselect(
                    "Selecione as NFs liquidadas:",
                    options=lista_nfs_7,
                    key="ms_liquidar_nf"
                )

                if st.button("💎 Confirmar Liquidação das NFs", use_container_width=True, type="primary"):
                    if nfs_para_liquidar:
                        with st.spinner("Processando liquidação..."):
                            count_nups = 0
                            for nf_sel in nfs_para_liquidar:
                                # Busca todos os NUPs vinculados a essa NF específica
                                nups_vinculados = df_status_7[df_status_7['nf'] == nf_sel]['nup'].tolist()
                                for nup in nups_vinculados:
                                    mover_status(nup, 8)
                                    dados_f = df_status_7[df_status_7['nup'] == nup].iloc[0]
                                    registrar_acao(nup, dados_f['Numero_da_fatura'], "LIQUIDACAO_EFETUADA", f"Liquidada via NF: {nf_sel}")
                                    registrar_historico(nup, dados_f['Numero_da_fatura'], "7", "8", dados_f['valor_apresentado'], f"Liquidação via NF {nf_sel}")
                                    count_nups += 1
                            
                            st.success(f"✅ {len(nfs_para_liquidar)} NFs processadas ({count_nups} faturas movidas)!")
                            time.sleep(1.2)
                            st.rerun()
                    else:
                        st.warning("Selecione ao menos uma NF.")

                st.dataframe(
                    df_status_7[['nf', 'ose', 'nup', 'Numero_da_fatura', 'valor_liquido', 'mes_sigla']],
                    use_container_width=True, hide_index=True
                )
            else:
                st.info("Não há faturas aguardando liquidação.")

            st.divider()

            # =================================================================
            # SEÇÃO 2: PAGAMENTO FINAL (Status 8 -> 9)
            # =================================================================
            st.markdown("### 💰 2. Efetuar Pagamento (por NF)")
            st.write("Conclua o processo preenchendo a OB e selecionando as NFs.")

            df_status_8 = df[df['status'] == 8].copy()

            if not df_status_8.empty:
                df_status_8['mes_sigla'] = df_status_8['mes_competencia'].map(mapa_meses)
                lista_nfs_8 = df_status_8['nf'].dropna().unique().tolist()

                # --- 🛠️ NOVA LÓGICA DE MÁSCARA DA OB ---
                ano_atual = datetime.now().year
                prefixo_ob = f"78770000001{ano_atual}OB"

                col_ob1, col_ob2 = st.columns([2, 1])
                with col_ob1:
                    # Mostra o prefixo fixo para o militar não precisar digitar
                    st.info(f"Prefixo Automático: **{prefixo_ob}**")
                with col_ob2:
                    # Campo para os 6 dígitos finais
                    ob_6_digitos = st.text_input("6 Dígitos Finais:", max_chars=6, placeholder="800116", key="ob_final_6")

                # Monta a OB completa para gravar na planilha
                ob_completa = prefixo_ob + ob_6_digitos

                nfs_para_pagar = st.multiselect(
                    "Selecione as NFs pagas:",
                    options=lista_nfs_8,
                    key="ms_pagar_nf"
                )

                if st.button("🚀 Confirmar Pagamento das NFs", use_container_width=True, type="primary"):
                    # Validação rigorosa dos 6 dígitos
                    if not ob_6_digitos or len(ob_6_digitos) != 6 or not ob_6_digitos.isdigit():
                        st.error(f"⚠️ Erro: Informe exatamente os 6 dígitos finais da OB. (Ex: {prefixo_ob}**800116**)")
                    elif not nfs_para_pagar:
                        st.warning("Selecione ao menos uma NF.")
                    else:
                        with st.spinner("Finalizando pagamentos e registrando OB..."):
                            count_nups_pago = 0
                            aba_proc = sh.worksheet("SISAFA-NAVAL-processos")
                            
                            for nf_sel in nfs_para_pagar:
                                df_vinculados = df_status_8[df_status_8['nf'] == nf_sel]
                                
                                for _, row_fatura in df_vinculados.iterrows():
                                    nup = row_fatura['nup']
                                    
                                    # 1. Atualiza o Status para 9
                                    mover_status(nup, 9)
                                    
                                    # 2. Grava a OB COMPLETA na coluna 17
                                    try:
                                        celula = aba_proc.find(str(nup))
                                        if celula:
                                            # Aqui gravamos a variável 'ob_completa'
                                            aba_proc.update_cell(celula.row, 17, ob_completa) 
                                    except:
                                        pass 
                                    
                                    # 3. Registra nos Logs e Histórico com a OB COMPLETA
                                    fatura_n = row_fatura['Numero_da_fatura']
                                    registrar_acao(nup, fatura_n, "PAGAMENTO_EFETUADO", f"Pago via NF: {nf_sel} | OB: {ob_completa}")
                                    registrar_historico(nup, fatura_n, "8", "9", row_fatura['valor_apresentado'], f"Pagamento via NF {nf_sel} - OB {ob_completa}")
                                    count_nups_pago += 1
                            
                            st.success(f"🎊 {len(nfs_para_pagar)} NFs pagas! OB {ob_completa} registrada.")
                            time.sleep(1.2)
                            st.rerun()

                st.dataframe(
                    df_status_8[['nf', 'ose', 'nup', 'Numero_da_fatura', 'valor_liquido', 'mes_sigla']],
                    use_container_width=True, hide_index=True
                )
            else:
                st.info("Não há faturas aguardando confirmação de pagamento.")

        # --- ABA 4: ESTATÍSTICAS E INDICADORES (KPIs Financeiros) ---
        with tab4:
        
            # ==========================================
            # 🎨 ESTILIZAÇÃO CSS (Neon no Fundo Branco)
            # ==========================================
            st.markdown("""
            <style>
            /* Cards base (Efeito Neon Suave no Branco) */
            .neon-card {
                background: #ffffff;
                border-radius: 12px;
                padding: 20px;
                margin-bottom: 20px;
                text-align: center;
                box-shadow: 0 4px 15px rgba(0, 0, 0, 0.05);
                transition: transform 0.2s;
                border-bottom: 4px solid #eee;
            }
            .neon-card:hover { transform: translateY(-3px); }
            
            /* Títulos dos Cards */
            .nc-title { font-size: 0.9rem; font-weight: 700; color: #555; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 10px; }
            .nc-value { font-size: 2.2rem; font-weight: 900; margin-bottom: 5px; }
            .nc-sub { font-size: 0.85rem; font-weight: 600; color: #777; }
            
            /* Cores Neon / Futuristas Específicas */
            .card-cyan { border-bottom-color: #00e5ff; box-shadow: 0 10px 20px rgba(0, 229, 255, 0.15); }
            .card-cyan .nc-value { color: #00b8d4; text-shadow: 0 0 10px rgba(0, 229, 255, 0.3); }
            
            .card-purple { border-bottom-color: #d500f9; box-shadow: 0 10px 20px rgba(213, 0, 249, 0.15); }
            .card-purple .nc-value { color: #aa00ff; text-shadow: 0 0 10px rgba(213, 0, 249, 0.3); }
            
            .card-alert { border-bottom-color: #ff1744; box-shadow: 0 10px 20px rgba(255, 23, 68, 0.15); background: #fffcfc;}
            .card-alert .nc-value { color: #d50000; text-shadow: 0 0 10px rgba(255, 23, 68, 0.3); }
            
            .card-green { border-bottom-color: #00e676; box-shadow: 0 10px 20px rgba(0, 230, 118, 0.15); }
            .card-green .nc-value { color: #00c853; text-shadow: 0 0 10px rgba(0, 230, 118, 0.3); }
            
            .card-orange { border-bottom-color: #ff9100; box-shadow: 0 10px 20px rgba(255, 145, 0, 0.15); }
            .card-orange .nc-value { color: #ff6d00; text-shadow: 0 0 10px rgba(255, 145, 0, 0.3); }
            </style>
            """, unsafe_allow_html=True)

            st.markdown("""
                        <div style="text-align: center; padding: 30px; border-top: 2px solid #2e6b54; margin-top: 40px; background-color: rgba(46, 107, 84, 0.05); border-radius: 0 0 15px 15px;">
                            <p style="
                                color: #2e6b54; 
                                font-weight: 900; 
                                font-size: 1.8rem; 
                                letter-spacing: 3px; 
                                line-height: 1.2;
                                text-shadow: 0 0 10px #2e6b54, 0 0 20px #2e6b54, 0 0 30px #2e6b54;
                            ">
                                "RESTARÁ SEMPRE MUITO O QUE FAZER"
                            </p>
                            <p style="
                                color: #555; 
                                font-size: 1.1rem; 
                                font-weight: 700; 
                                margin-top: -10px;
                                text-transform: uppercase;
                                letter-spacing: 1px;
                            ">
                                (SEPÚLVEDA, A.C.M)
                            </p>
                        </div>
                        """, unsafe_allow_html=True)

            st.header("⚡ Painel de Execução Financeira")
            st.info("Monitoramento em tempo real do fluxo de empenho, liquidação e pagamento.")

            # --- 1. PREPARAÇÃO DOS DADOS GERAIS ---
            df_e = df.copy()
            df_e['v_ap_num'] = df_e['valor_apresentado'].apply(limpar_valor)
            df_e['v_liq_num'] = df_e['valor_liquido'].apply(limpar_valor)
            
            # Preparação de Datas (Uso explícito do nome da coluna para evitar erros de índice)
            # Garantimos que, mesmo se a ordem das colunas mudar na planilha, o cálculo permanece correto
            # --- MOTOR DE NORMALIZAÇÃO DE DATAS ---
            def normalizar_data(valor):
                v_str = str(valor).strip()
                if not v_str or v_str.lower() == 'nan' or v_str == 'None':
                    return pd.NaT
                
                # Tenta converter formatos mistos (ISO e BR)
                # dayfirst=True ajuda no formato 12/03/2026, mas o formato ISO (2026-05-26) é autodetectado
                return pd.to_datetime(v_str, dayfirst=True, errors='coerce')

            # Aplica a normalização na coluna 14 (índice 13)
            df_e['dt_mov'] = df_e.iloc[:, 13].apply(normalizar_data)

            # Cálculo de dias com tratamento para NaT (evita erro se a data ainda estiver inválida)
            hoje = pd.Timestamp.now()
            df_e['dias_na_fase'] = (hoje - df_e['dt_mov']).dt.days.fillna(0).astype(int)
            
            
            # Criação do campo "Competência Unificada" para os gráficos
            
            # Criação do campo "Competência Unificada" para os gráficos (com blindagem de colunas)
            col_mes = 'mes_sigla' if 'mes_sigla' in df_e.columns else 'mes_competencia'
            
            if col_mes in df_e.columns and 'ano_competencia' in df_e.columns:
                df_e['competencia_grafico'] = df_e[col_mes].astype(str) + "/" + df_e['ano_competencia'].astype(str)
            else:
                # Fallback de segurança caso as colunas tenham nomes diferentes na planilha
                df_e['competencia_grafico'] = "Sem Data"

            # --- FILTROS RÁPIDOS ---
            c_f1, c_f2 = st.columns(2)
            anos_disp = sorted(df_e['ano_competencia'].unique(), reverse=True)
            ano_sel = c_f1.selectbox("Filtrar Painel por Ano:", ["Todos"] + list(anos_disp), key="f_ano_exec_tatico")
            if ano_sel != "Todos": 
                df_e = df_e[df_e['ano_competencia'] == ano_sel]

            st.markdown("<br>", unsafe_allow_html=True)

            # ==========================================
            # 🚀 LINHA 1: A ENTRADA E O ALERTA MÁXIMO
            # ==========================================
            c1, c2, c3 = st.columns(3)
            
            # STATUS 3: Aguardando Recebimento na Execução
            df_s3 = df_e[df_e['status'] == 3]
            v_s3 = df_s3['v_liq_num'].sum()
            t_s3 = df_s3['dias_na_fase'].mean() if not df_s3.empty else 0
            
            c1.markdown(f"""
            <div class="neon-card card-cyan">
                <div class="nc-title">1. Na Porta da Execução (Auditados)</div>
                <div class="nc-value">{len(df_s3)} Faturas</div>
                <div class="nc-sub">Volume: R$ {v_s3:,.2f}</div>
                <div class="nc-sub">⏳ Média de {t_s3:.0f} dias aguardando envio</div>
            </div>
            """, unsafe_allow_html=True)

            # STATUS 4: Aguardando Emissão de NE
            df_s4 = df_e[df_e['status'] == 4]
            v_s4 = df_s4['v_liq_num'].sum()
            
            c2.markdown(f"""
            <div class="neon-card card-purple">
                <div class="nc-title">2. Aguardando Emissão de NE ⌛</div>
                <div class="nc-value">{len(df_s4)} Faturas</div>
                <div class="nc-sub">Volume: R$ {v_s4:,.2f}</div>
                <div class="nc-sub">Prontos para empenho</div>
            </div>
            """, unsafe_allow_html=True)

            # STATUS 5: Empenhados SEM Fiscalização (PONTO DE ATENÇÃO)
            df_s5 = df_e[df_e['status'] == 5]
            max_dias_s5 = df_s5['dias_na_fase'].max() if not df_s5.empty else 0
            
            c3.markdown(f"""
            <div class="neon-card card-alert">
                <div class="nc-title">⚠️ 3. Empenhados e não encaminhados à fiscalização de contratos</div>
                <div class="nc-value">{len(df_s5)} Faturas</div>
                <div class="nc-sub">Falta encaminhar para fiscalização</div>
                <div class="nc-sub" style="color:#d50000; font-weight:bold;">🔥 Processo mais antigo: {max_dias_s5:.0f} dias parado</div>
            </div>
            """, unsafe_allow_html=True)


            # ==========================================
            # 🚀 LINHA 2: O FLUXO FINAL
            # ==========================================
            c4, c5, c6 = st.columns(3)

            # STATUS 6: Aguardando Recebimento pelos Fiscais
            df_s6 = df_e[df_e['status'] == 6]
            c4.markdown(f"""
            <div class="neon-card card-orange">
                <div class="nc-title">4. Pendente de entrega dos Fiscais</div>
                <div class="nc-value">{len(df_s6)} Faturas</div>
                <div class="nc-sub">Caixa de entrada da fiscalização</div>
            </div>
            """, unsafe_allow_html=True)

            # STATUS 7: Em Liquidação
            df_s7 = df_e[df_e['status'] == 7]
            t_s7 = df_s7['dias_na_fase'].mean() if not df_s7.empty else 0
            c5.markdown(f"""
            <div class="neon-card card-cyan">
                <div class="nc-title">5. Em Liquidação (NF entregue, mas não liquidada!)</div>
                <div class="nc-value">{len(df_s7)} Faturas</div>
                <div class="nc-sub">⏳ Média de {t_s7:.0f} dias nesta fase</div>
            </div>
            """, unsafe_allow_html=True)

            # STATUS 8: Liquidado, Não Pago
            df_s8 = df_e[df_e['status'] == 8]
            v_s8 = df_s8['v_liq_num'].sum()
            c6.markdown(f"""
            <div class="neon-card card-green">
                <div class="nc-title">6. Pronto para Pagar</div>
                <div class="nc-value">{len(df_s8)} Faturas</div>
                <div class="nc-sub">Volume a pagar bruto: R$ {v_s8:,.2f}</div>
            </div>
            """, unsafe_allow_html=True)

            st.divider()


            # ==========================================
            # 📈 GRÁFICO ANALÍTICO GERAL (Foco em NE)
            # ==========================================
            import plotly.express as px

            # Status 4 (Aguard. NE) por Competência COM VALOR FINANCEIRO
            if not df_s4.empty:
                # Agrupa contando a quantidade E somando o valor líquido
                df_ne = df_s4.groupby(['ano_competencia', col_mes, 'competencia_grafico']).agg(
                    Qtd=('nup', 'count'),
                    Valor=('v_liq_num', 'sum')
                ).reset_index()
                
                # Converte para numérico temporariamente para ordenar a linha do tempo (de baixo para cima)
                df_ne['ano_num'] = pd.to_numeric(df_ne['ano_competencia'], errors='coerce')
                df_ne['mes_num'] = pd.to_numeric(df_ne[col_mes], errors='coerce')
                df_ne = df_ne.sort_values(by=['ano_num', 'mes_num'], ascending=True)

                # Cria o texto tecnológico para mostrar dentro da barra (Qtd + Valor)
                df_ne['texto_barra'] = df_ne.apply(lambda row: f"{row['Qtd']} faturas | R$ {row['Valor']:,.2f}", axis=1)

                fig2 = px.bar(
                    df_ne, x='Qtd', y='competencia_grafico', orientation='h',
                    title="📄 Aguard. Emissão de NE (Qtd e Valor Total)",
                    labels={'Qtd': 'Quantidade', 'competencia_grafico': 'Competência', 'texto_barra': 'Resumo'},
                    color_discrete_sequence=['#d500f9'],
                    template="plotly_white",
                    text='texto_barra',
                    hover_data={'Valor': ':,.2f'}
                )
                
                fig2.update_traces(
                    marker_line_color='#aa00ff', marker_line_width=1.5, opacity=0.8, 
                    textposition='inside',
                    textfont=dict(color='white', size=13, family='Arial')
                )
                
                fig2.update_layout(
                    plot_bgcolor="rgba(0,0,0,0)", margin=dict(t=50, l=10, r=10, b=10),
                    yaxis={'categoryorder': 'array', 'categoryarray': df_ne['competencia_grafico']}
                )
                st.plotly_chart(fig2, use_container_width=True)
            else:
                st.success("🎉 Não há processos aguardando Nota de Empenho!")

            # =================================================================
            # =================================================================
            # 🏢 PAINEL DE ANÁLISE FINANCEIRA POR EMPRESA (TÁTICO)
            # =================================================================
            st.divider()
            st.header("🏢 Painel de Análise Financeira por Empresa (OSE)")
            
            # Pega a lista de todas as empresas únicas no radar da execução
            lista_empresas = sorted(df_e['ose'].dropna().unique())
            ose_sel = st.selectbox("Selecione a Organização para gerar o Dossiê:", [""] + lista_empresas, key="sel_ose_execucao")

            if ose_sel:
                # Filtra o dataframe apenas para a OSE selecionada
                df_ose_exec = df_e[df_e['ose'] == ose_sel].copy()
                
                # Mapeamento oficial dos status
                mapa_status_nomes = {                
                    1: "1. 📥 Cadastrada", 2: "2. 🩺 Em Auditagem", 3: "3. ✅ Auditada (Aguard. Execução)",
                    4: "4. 💰 Aguardando NE", 5: "5. 🏦 Empenhada", 6: "6. 📝 Aguardando Fiscal",
                    7: "7. ⏳ Em liquidação", 8: "8. 🖥️ Liquidada", 9: "9. 💸 Paga"
                }
                df_ose_exec['etapa_nome'] = df_ose_exec['status'].map(mapa_status_nomes)

                c_metric1, c_metric2 = st.columns(2)
                
                # --- CORREÇÃO: TOTAL GLOBAL SEM FILTRO DE STATUS ---
                volume_total_ose = df_ose_exec['v_liq_num'].sum()
                qtd_total_ose = len(df_ose_exec)
                
                c_metric1.markdown(f"""
                <div class="neon-card card-cyan" style="padding:15px; margin-top:10px;">
                    <div class="nc-title">Volume Financeiro Total</div>
                    <div class="nc-value" style="font-size:1.8rem;">R$ {volume_total_ose:,.2f}</div>
                </div>
                """, unsafe_allow_html=True)
                
                c_metric2.markdown(f"""
                <div class="neon-card card-purple" style="padding:15px; margin-top:10px;">
                    <div class="nc-title">Faturas Cadastradas</div>
                    <div class="nc-value" style="font-size:1.8rem;">{qtd_total_ose} faturas</div>
                </div>
                """, unsafe_allow_html=True)

                # --- GRÁFICO DE PIZZA (Montante Financeiro por Etapa - TODOS OS STATUS) ---
                st.markdown("#### 🍩 Distribuição do Montante Financeiro por Etapa")
                df_pizza = df_ose_exec.groupby('etapa_nome')['v_liq_num'].sum().reset_index()
                
                if not df_pizza.empty:
                    fig_pie = px.pie(
                        df_pizza, values='v_liq_num', names='etapa_nome',
                        hole=0.4,
                        # Adicionadas mais cores à paleta para acomodar todos os 9 status sem repetir
                        color_discrete_sequence=['#00e5ff', '#d500f9', '#00e676', '#ff9100', '#ff1744', '#2979ff', '#00bfa5', '#ffd600', '#c51162'],
                        template="plotly_white"
                    )
                    # Formata para aparecer o valor em R$ no gráfico
                    fig_pie.update_traces(
                        textposition='outside', 
                        texttemplate='<b>%{label}</b><br>R$ %{value:,.2f} (%{percent})',
                        marker=dict(line=dict(color='#ffffff', width=2))
                    )
                    fig_pie.update_layout(height=450, showlegend=False, margin=dict(l=10, r=10, t=30, b=10))
                    st.plotly_chart(fig_pie, use_container_width=True)
                else:
                    st.info("Não há volume financeiro pendente para esta OSE.")

                # =================================================================
                # 📧 COMUNICAÇÃO E RELATÓRIOS OFICIAIS PARA O FORNECEDOR
                # =================================================================
                st.markdown("#### ✉️ Comunicação com a OSE")
                st.info("Gere o texto padrão atualizado para responder e-mails ou emita o Dossiê em PDF.")

                col_btn1, col_btn2 = st.columns(2)

                # BOTÃO 1: GERAR TEXTO PARA E-MAIL/WHATSAPP
                # (Se não tiver estes imports no topo do arquivo, deixe-os aqui alinhados à esquerda)
                import base64
                import gc
                import matplotlib.pyplot as plt

                # --- BOTÃO 1: GERAR TEXTO PARA E-MAIL/WHATSAPP ---
                with col_btn1:
                    if st.button("📑 Gerar texto de Panorama", use_container_width=True):
                        with st.spinner("Sincronizando dados..."):
                            
                            explica_etapa = {
                                1: "Registro e conferência inicial da documentação em nossa Secretaria.",
                                2: "Divisão de Auditoria recebeu as faturas e iniciou a análise técnica detalhada dos serviços e materiais cobrados.",
                                3: "Auditoria técnica concluída com sucesso e aguardando o recebimento no setor financeiro.",
                                4: "Aguardando a reserva orçamentária, com vistas à emissão da Nota de Empenho.",
                                5: "Recurso orçamentário já reservado especificamente para estas faturas. Oportunamente, destaca-se que serão envidados os esforços necessários para o encaminhamento da respectiva NE com a maior celeridade possível!",
                                6: "Fase em que o fiscal do contrato irá apreciar a Nota de Empenho, bem como realizará o devido contato com a empresa, com vistas à emissão dos documentos fiscais pertinentes.",
                                7: "Fase em que a empresa deve emitir e enviar a Nota Fiscal para o Hospital. A Nota Fiscal é certificada pelo gestor e, posteriormente, encaminhada para a Seção de Execução Financeira.",
                                8: "Liquidação da Nota Fiscal no Sistema Integrado de Administração Financeira do Governo Federal (SIAFI) 🖥️, conforme o estabelecido no artigo 63 da Lei 4.320/64.",
                                9: "Pagamento autorizado pelo Ordenador de Despesas. Nessa etapa, o pagamento demora, em média, um dia útil após a aprovação para ser creditado em conta-corrente. Outrossim, salienta-se que, por ocasião dos pagamentos, são realizados os abatimentos tributários devidos."
                            }

                            meses_map = {1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR", 5: "MAI", 6: "JUN", 7: "JUL", 8: "AGO", 9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ"}
                            df_ose_exec['dt_ordem'] = pd.to_datetime(df_ose_exec['dt_mov'], errors='coerce')
                            resumo_corpo = ""
                            etapas_ativas = sorted(df_ose_exec['status'].unique())
                            
                            for st_id in etapas_ativas:
                                if st_id >= 9: continue
                                df_etapa = df_ose_exec[df_ose_exec['status'] == st_id].copy()
                                if not df_etapa.empty:
                                    nome_da_etapa = mapa_status_nomes.get(st_id, f"Etapa {st_id}")
                                    total_etapa = df_etapa['v_liq_num'].sum()
                                    descricao = explica_etapa.get(st_id, "")
                                    resumo_corpo += f"\n🔹 **{nome_da_etapa.upper()}**\n"
                                    resumo_corpo += f"   - {descricao}\n"
                                    resumo_corpo += f"   - Volume Total na Etapa: R$ {total_etapa:,.2f}\n"
                                    
                                    competencias = df_etapa['competencia_grafico'].unique()
                                    for comp in competencias:
                                        df_comp = df_etapa[df_etapa['competencia_grafico'] == comp]
                                        lista_faturas = ", ".join(df_comp['Numero_da_fatura'].astype(str).tolist())
                                        subtotal_comp = df_comp['v_liq_num'].sum()
                                        try:
                                            mes_num, ano_str = comp.split('/')
                                            mes_sigla = meses_map.get(int(mes_num), mes_num)
                                            rotulo_entrada = f"{mes_sigla}/{ano_str}"
                                        except:
                                            rotulo_entrada = comp
                                        resumo_corpo += f"     ➔ Competência {rotulo_entrada}: Faturas [{lista_faturas}] — Subtotal: R$ {subtotal_comp:,.2f}\n"
                                    resumo_corpo += "───────────────────────────────────────\n"

                            msg_final = f"""Prezado (a) representante da empresa {ose_sel},

Cumprimentando-o (a) cordialmente, seguem abaixo algumas orientações, bem como um panorama atualizado de seus processos ora em trâmite no Hospital Naval de Brasília (HNBra):

📌 (i) DO CRITÉRIO DE PAGAMENTO
Este hospital realiza a emissão das Notas de Empenho em estrita ordem cronológica, mediante disponibilidade orçamentária, a partir da data da entrada da(s) fatura(s) em nossa Secretaria. Nesse contexto, é útil uma análise literal da Lei 14.133/21 (Art. 141): 
                        
"No dever de pagamento pela Administração, será observada a ordem cronológica para cada fonte diferenciada de recursos [...]"

📊 (ii) COMPOSIÇÃO ATUAL DOS PROCESSOS
{resumo_corpo}
🚀 (iii) PRÓXIMOS PASSOS
As faturas supracitadas seguem em fluxo contínuo de processamento. Assim que concluídas as etapas de conferência e reserva orçamentária, as Notas de Empenho correspondentes a cada fatura serão encaminhadas para a respectiva emissão de Nota Fiscal e posterior liquidação.

Reitera-se nosso compromisso com a transparência e eficiência em nossos atos administrativos. 

Estamos à disposição para eventuais esclarecimentos. Gratos pela distinta parceria! 🤝⚓🇧🇷 

Cordialmente,
"""
                            # Salva na memória do Streamlit
                            st.session_state['panorama_gerado_exec'] = msg_final

                # --- BOTÃO 2: GERAR E BAIXAR O PDF OFICIAL ---
                with col_btn2:
                    if st.button("🖨️ Processar Dossiê em PDF", use_container_width=True):
                        with st.spinner("Compilando dados e gráficos (aguarde)..."):
                            try:
                                # 1. Gera o PDF
                                pdf_bytes = gerar_relatorio_ose_pdf(ose_sel, df_ose_exec, volume_total_ose, qtd_total_ose, fig_pie)
                                
                                # 2. Desafoga a memória RAM do servidor
                                plt.close('all')
                                gc.collect()
                                
                                # 3. Converte a fundo para Base64
                                b64 = base64.b64encode(pdf_bytes).decode()
                                nome_arquivo = f"Dossie_{ose_sel.replace(' ', '_')}.pdf"
                                
                                # 4. Cria o Botão HTML encapsulado em uma <div> (Isto impede o Streamlit de ler como texto)
                                html_button = (
                                    f'<div style="text-align: center;">'
                                    f'<a href="data:application/pdf;base64,{b64}" download="{nome_arquivo}" '
                                    f'style="display: block; padding: 12px; background-color: #00E676; color: #1e1e1e; '
                                    f'text-align: center; text-decoration: none; font-size: 16px; font-family: sans-serif; '
                                    f'border-radius: 8px; font-weight: bold; width: 100%; box-sizing: border-box;">'
                                    f'📥 PDF PRONTO! CLIQUE AQUI PARA BAIXAR'
                                    f'</a>'
                                    f'</div>'
                                )
                                
                                # 5. Renderiza a estrutura blindada
                                st.markdown(html_button, unsafe_allow_html=True)
                                
                            except Exception as e:
                                st.error(f"Erro ao compilar o PDF: {e}")

                # --- EXIBIÇÃO DO PANORAMA GERADO ---
                # TOTALMENTE FORA DAS COLUNAS (Alinhado com a declaração do "with col_btn1" / "with col_btn2")
                if 'panorama_gerado_exec' in st.session_state:
                    st.success("✅ Panorama gerado com sucesso! Copie o texto abaixo:")
                    st.text_area(
                        "Texto para E-mail / WhatsApp", 
                        st.session_state['panorama_gerado_exec'], 
                        height=350
                    )     



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

        # --- ABA 6: RELACIONAMENTO (Módulo Execução Financeira) ---
        with tab6:
            st.subheader("🤝 Central de Relacionamento (Setor: FINANCEIRO)")
            st.write("Dúvidas financeiras e questionamentos de faturas enviados pelas OSEs.")

            try:
                # 1. Carregamos as mensagens da aba correspondente
                aba_msg = sh.worksheet(ABA_MENSAGENS)
                dados_brutos = aba_msg.get_all_records()
                df_msg = pd.DataFrame(dados_brutos)
                
                if df_msg.empty:
                    st.info("Nenhuma mensagem registrada no sistema.")
                else:
                    # --- FILTRO POR SETOR (FINANCEIRO) ---
                    df_msg_exec = df_msg[df_msg['setor_destino'] == "FINANCEIRO"].copy()

                    if df_msg_exec.empty:
                        st.info("✅ Tudo em dia! Nenhuma mensagem pendente para o setor Financeiro.")
                    else:
                        # 2. Métricas do Setor
                        pendentes = len(df_msg_exec[df_msg_exec['status_msg'] == 'PENDENTE'])
                        c1, c2 = st.columns(2)
                        c1.metric("Mensagens Recebidas", len(df_msg_exec))
                        c2.metric("📩 Pendentes", pendentes, delta_color="inverse")
                        
                        st.divider()

                        # 3. Inbox de Mensagens (Usando os nomes das 10 colunas da sua planilha)
                        st.write("**📥 Inbox da Execução Financeira:**")
                        
                        # Ordenamos para ver o que é PENDENTE primeiro
                        df_exibir = df_msg_exec.sort_values(by='status_msg', ascending=False)
                        
                        st.dataframe(
                            df_exibir[['Numero_da_fatura', 'nup', 'remetente', 'data_envio', 'status_msg']], 
                            use_container_width=True,
                            hide_index=True
                        )

                        st.markdown("---")

                        # 4. Área de Resposta do Financeiro
                        st.markdown("### ✍️ Dar Parecer Financeiro")
                        
                        # Criamos um rótulo amigável para seleção
                        df_msg_exec['label_selecao'] = (
                            "Fatura: " + df_msg_exec['Numero_da_fatura'].astype(str) + 
                            " | ID: " + df_msg_exec['id_mensagem'].astype(str)
                        )
                        
                        selecao_msg = st.selectbox("Selecione a mensagem para responder:", [""] + df_msg_exec['label_selecao'].tolist(), key="sel_msg_exec")

                        if selecao_msg:
                            # Localizamos os dados da mensagem alvo
                            item = df_msg_exec[df_msg_exec['label_selecao'] == selecao_msg].iloc[0]
                            id_msg_alvo = str(item['id_mensagem'])
                            
                            with st.container(border=True):
                                st.write(f"🏢 **OSE (Remetente):** {item['remetente']}")
                                st.write(f"📑 **NUP:** {item['nup']}")
                                st.chat_message("user").write(f"**Dúvida da OSE:**\n\n{item['texto']}")
                                
                                resp_exec = st.text_area("Parecer da Execução Financeira:", height=150, placeholder="Digite a resposta oficial...")
                                
                                if st.button("📤 ENVIAR PARECER FINANCEIRO", use_container_width=True):
                                    if resp_exec:
                                        with st.spinner("Gravando resposta..."):
                                            # AÇÃO 1: Log na Tabela-B
                                            registrar_acao(item['nup'], item['Numero_da_fatura'], "RESPOSTA_FINANCEIRA", f"Financeiro respondeu ID {id_msg_alvo}")
                                            
                                            # AÇÃO 2: Atualizar a linha na ABA_MENSAGENS (As 10 Colunas)
                                            try:
                                                # Localizamos a linha exata pelo ID Único (Coluna A)
                                                celula = aba_msg.find(id_msg_alvo)
                                                linha_idx = celula.row
                                                
                                                # Atualizamos: data_resposta (8), status_msg (9), respondido_por_nip (10)
                                                agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                                aba_msg.update_cell(linha_idx, 8, agora)
                                                aba_msg.update_cell(linha_idx, 9, "RESPONDIDO")
                                                aba_msg.update_cell(linha_idx, 10, str(st.session_state.user_id)) # NIP do militar

                                                st.success("Resposta enviada para o portal da OSE!")
                                                time.sleep(1.5)
                                                st.rerun()
                                            except Exception as err:
                                                st.error(f"Erro ao localizar mensagem na planilha: {err}")
                                    else:
                                        st.warning("Por favor, preencha o parecer antes de enviar.")

            except Exception as e:
                st.error(f"Erro na aba relacionamento do financeiro: {e}")

    # =================================================================
    # MÓDULO 4: FISCALIZAÇÃO DE CONTRATOS (Ajustado para FISCAL/FISCAL_GLOBAL)
    # =================================================================
    elif "FISCAL" in st.session_state.modulo_ativo:
        st.header("📋 Fiscalização de Contratos (OSE)")
        
        # --- 1. DEFINIÇÃO DOS MAPAS (A BASE DO TREINO) ---
        # Definir aqui garante que o NameError não apareça neste módulo
        mapa_status_fisc = {
            1: "1 - FATURA CADASTRADA", 2: "2 - EM AUDITAGEM", 3: "3 - AUDITADA",
            4: "4 - AGUARDANDO EMISSÃO DE NE", 5: "5 - FATURA EMPENHADA",
            6: "6 - AGUARDANDO EMISSÃO DE NF", 7: "7 - EM LIQUIDAÇÃO",
            8: "8 - FATURA LIQUIDADA", 9: "9 - FATURA PAGA"
        }

        cores_map = {
            "1 - FATURA CADASTRADA": "#95a5a6", "2 - EM AUDITAGEM": "#f39c12",
            "3 - AUDITADA": "#3498db", "4 - AGUARDANDO EMISSÃO DE NE": "#f1c40f",
            "5 - FATURA EMPENHADA": "#9b59b6", "6 - AGUARDANDO EMISSÃO DE NF": "#e67e22",
            "7 - EM LIQUIDAÇÃO": "#e74c3c", "8 - FATURA LIQUIDADA": "#1abc9c",
            "9 - FATURA PAGA": "#27ae60"
        }

        # --- 2. PREPARAÇÃO DE DADOS (VACINA DO CNPJ E NIP) ---
        df_tabela_a = carregar_dados_cache(ABA_TABELA_A)
        df_tabela_a.columns = [c.strip().replace(' ', '_').upper() for c in df_tabela_a.columns]
        
        # 2. Criamos o CNPJ_LIMPO de forma simples e segura
        df_tabela_a = df_tabela_a.loc[:, ~df_tabela_a.columns.duplicated()]

        # 3. Agora criamos o CNPJ_LIMPO com total segurança
        if 'CNPJ' in df_tabela_a.columns:
            # Garantimos que a coluna seja tratada como texto e limpamos
            df_tabela_a['CNPJ_LIMPO'] = (
                df_tabela_a['CNPJ']
                .astype(str)
                .str.replace(r'\.0$', '', regex=True) # Tira o .0 do Excel
                .str.replace(r'\D', '', regex=True)    # Tira pontos, traços e barras
                .str.zfill(14)                         # Garante os 14 dígitos
            )
        else:
            st.error("❌ Coluna 'CNPJ' não localizada. Verifique o cabeçalho da Tabela A.")
            st.stop()
                
        user_nip = str(st.session_state.user_id).strip().zfill(8)
        is_global = (user_nip == "95039023")

        if is_global:
            df_fiscal = df_tabela_a.copy()
        else:
            col_nip = "NIP_DO_GESTOR_TITULAR"
            col_nip_sub = "NIP_DO_GESTOR_SUBSTITUTO"
            for c in [col_nip, col_nip_sub]:
                if c in df_tabela_a.columns:
                    df_tabela_a[c] = df_tabela_a[c].apply(lambda x: str(x).split('.')[0].strip().zfill(8) if x else "")
            
            filtro = (df_tabela_a[col_nip] == user_nip) | (df_tabela_a[col_nip_sub] == user_nip)
            df_fiscal = df_tabela_a[filtro].copy()

        # --- 3. DEFINIÇÃO DAS ABAS ---
        tab_visao, tab_nf, tab_rel = st.tabs(["🔭 Visão Geral", "🧾 Empenhos aguardando NF", "💬 Relacionamento"])

        with tab_visao:
            st.subheader("Meus contratos")
            if df_fiscal.empty:
                st.warning(f"⚠️ Nenhum contrato vinculado ao NIP {user_nip}.")
            else:
                st.markdown(
                    f"Contratos sob sua responsabilidade (<span style='color: #2e6b54; font-weight: bold;'>Titular</span> ou <span style='color: #cba30c; font-weight: bold;'>Substituto</span>):", 
                    unsafe_allow_html=True
                )

                def style_rows(row):
                    if row["NIP_DO_GESTOR_TITULAR"] == user_nip:
                        return ['background-color: #2e6b54; color: white'] * len(row)
                    return ['background-color: #cba30c; color: black'] * len(row)

                st.dataframe(
                    df_fiscal[['CNPJ', 'RAZÃO_SOCIAL', 'NIP_DO_GESTOR_TITULAR', 'NIP_DO_GESTOR_SUBSTITUTO']]
                    .style.apply(style_rows, axis=1),
                    column_order=("CNPJ", "RAZÃO_SOCIAL"),
                    use_container_width=True, hide_index=True
                )

                st.divider()
                st.subheader("Situação geral")
                ose_sel = st.selectbox("Selecione a Organização:", [""] + df_fiscal['RAZÃO_SOCIAL'].tolist(), key="fisc_sel_final_v5")

                if ose_sel:
                    # Filtro de processos da OSE escolhida
                    cnpj_alvo = df_fiscal[df_fiscal['RAZÃO_SOCIAL'] == ose_sel]['CNPJ_LIMPO'].iloc[0]
                    
                    # Criamos a coluna vacinada no df principal para o match
                    df['cnpj_vacinado'] = df['cnpj'].astype(str).str.split('.').str[0].str.strip().str.zfill(14)
                    df_proc_fisc = df[df['cnpj_vacinado'] == cnpj_alvo].copy()

                    if not df_proc_fisc.empty:
                        df_proc_fisc['situação_texto'] = df_proc_fisc['status'].map(mapa_status_fisc)
                        
                        # --- 1. TABELA EM CIMA ---
                        st.write(f"📋 **Processos de {ose_sel}:**")
                        st.dataframe(
                            df_proc_fisc[['nup', 'Numero_da_fatura', 'situação_texto']].rename(columns={'situação_texto': 'Situação'}), 
                            use_container_width=True, 
                            hide_index=True
                        )

                        st.divider() # Uma linha para separar

                        # =================================================================
                        # 📊 RESUMO GERENCIAL 
                        # =================================================================
                        st.write("") 
                        st.header("📊 Resumo Gerencial da Fiscalização")
                        
                        # 1. Preparação de Dados
                        df_proc_fisc['v_liq'] = df_proc_fisc['valor_liquido'].apply(limpar_valor)
                        
                        mapa_status_nomes = {                
                            1: "1. 📥 Cadastrada", 2: "2. 🩺 Em Auditagem", 3: "3. ✅ Auditada",
                            4: "4. 💰 Aguardando NE", 5: "5. 🏦 Empenhada", 6: "6. 📝 Aguardando NF",
                            7: "7. ⏳ Em liquidação", 8: "8. 🖥️ Liquidada", 9: "9. 💸 Paga"
                        }
                        df_proc_fisc['etapa_nome'] = df_proc_fisc['status'].map(mapa_status_nomes)

                        # --- SEÇÃO 1: MÉTRICAS ---
                        c1, c2 = st.columns(2)
                        with c1:
                            tramito = df_proc_fisc[df_proc_fisc['status'] < 9]['v_liq'].sum()
                            st.metric("💰 Volume Financeiro em Trâmite", f"R$ {tramito:,.2f}")
                        with c2:
                            qtd_ativos = len(df_proc_fisc[df_proc_fisc['status'] < 9])
                            st.metric("📑 Total de Processos Ativos", f"{qtd_ativos} faturas")

                        st.divider()

                        # --- SEÇÃO 2: EQUILÍBRIO DO FLUXO (PIZZA) ---
                        st.markdown("#### 📚 Composição das faturas (por etapa)")
                        df_pizza = df_proc_fisc['etapa_nome'].value_counts().reset_index()
                        
                        fig_pie = px.pie(
                            df_pizza, values='count', names='etapa_nome',
                            hole=0.5,
                            color_discrete_sequence=px.colors.sequential.Mint_r
                        )
                        fig_pie.update_traces(textposition='outside', textinfo='percent+label')
                        fig_pie.update_layout(height=400, showlegend=False, margin=dict(l=10, r=10, t=30, b=10))
                        st.plotly_chart(fig_pie, use_container_width=True)

                        # --- SEÇÃO 3: DISTRIBUIÇÃO FINANCEIRA (BARRAS INVERTIDAS) ---
                        st.markdown("#### 💵 Distribuição Financeira por etapa")
                        
                        df_financeiro = df_proc_fisc.groupby('etapa_nome')['v_liq'].sum().reset_index()
                        
                        # --- AQUI ESTÁ A MUDANÇA: ascending=False para 1 ficar no TOPO ---
                        df_financeiro = df_financeiro.sort_values('etapa_nome', ascending=False)
                        
                        fig_bar = px.bar(
                            df_financeiro, 
                            x='v_liq', 
                            y='etapa_nome',
                            orientation='h',
                            color='v_liq',
                            color_continuous_scale=['#D1F2EB', '#76D7C4', '#1ABC9C', '#148F77'],
                            template="plotly_white"
                        )
                        
                        fig_bar.update_traces(
                            texttemplate='R$ %{x:,.2f}',
                            textposition='outside',
                            marker_line_color='#0E6251',
                            marker_line_width=1, 
                            opacity=0.9
                        )
                        
                        fig_bar.update_layout(
                            height=500,
                            margin=dict(l=0, r=100, t=20, b=20), # Margem direita extra para o valor real
                            xaxis_title="Volume Total (R$)",
                            yaxis_title="",
                            showlegend=False,
                            coloraxis_showscale=False
                        )
                        st.plotly_chart(fig_bar, use_container_width=True)

                        st.caption(f"🕒 Dados sincronizados em: {pd.Timestamp.now().strftime('%d/%m/%Y %H:%M')}")

        
                        # =================================================================
                        # =================================================================
                        # 📧 SEÇÃO: GERADOR DE PANORAMA PARA FORNECEDOR (VERSÃO FINALÍSSIMA)
                        # =================================================================
                        st.divider()
                        st.subheader("✉️ Comunicação com o Fornecedor")

                        # 1. Identificação Automática do Gestor (Assinatura)
                        nome_gestor_auto = st.session_state.get('usuario_nome', 'Auditor Responsável')

                        if st.button("📑 Gerar Panorama de Pagamento", use_container_width=True, key="btn_panorama_v3"):
                            with st.spinner("⚓ Sincronizando dados e formatando panorama oficial..."):
                                
                                # 2. Identificação da Coluna de Entrada (Competência/Mês)
                                col_comp_local = None
                                for c in df_proc_fisc.columns:
                                    if "competencia" in c.lower().strip() or "comp" in c.lower().strip():
                                        col_comp_local = c
                                        break
                                
                                if not col_comp_local:
                                    st.error("❌ Coluna de competência não localizada.")
                                else:
                                    # --- CORREÇÃO DO ERRO (CRIAÇÃO DA COLUNA DT_ORDEM) ---
                                    def tratar_data_panorama(v):
                                        v_str = str(v).strip().replace('.0', '')
                                        if not v_str or v_str == 'nan': return pd.NaT
                                        # Se for apenas número (1 a 12), assume o ano atual
                                        if v_str.isdigit() and 1 <= int(v_str) <= 12:
                                            return pd.to_datetime(f"{v_str.zfill(2)}/2026", format='%m/%Y')
                                        return pd.to_datetime(v_str, errors='coerce', dayfirst=True)

                                    # Criamos a coluna necessária para a ordenação antes do loop
                                    df_proc_fisc['dt_ordem'] = df_proc_fisc[col_comp_local].apply(tratar_data_panorama)

                                    # 3. Dicionário de Explicações Didáticas para o Fornecedor (Mantido seu texto)
                                    explica_etapa = {
                                        1: "Registro e conferência inicial da documentação em nossa Secretaria.",
                                        2: "Divisão de Auditoria recebeu as faturas e iniciou a análise técnica detalhada dos serviços e materiais cobrados, com vistas à emissão de glosas, controle de custos, entre outros serviços.",
                                        3: "Auditoria técnica concluída com sucesso e aguardando o envio para o setor financeiro.",
                                        4: "Aguardando a reserva orçamentária, com vistas à emissão da Nota de Empenho.",
                                        5: "Recurso orçamentário já reservado especificamente para estas faturas. Oportunamente, destaca-se que serão envidados os esforços necessários para o encaminhamento da respectiva NE com a maior celeridade possível!",
                                        6: "Fase em que o fiscal do contrato irá apreciar a Nota de Empenho, bem como realizará o devido contato com a empresa, com vistas à emissão do documentos fiscais pertinentes.",
                                        7: "Fase em que a empresa deve emitir e enviar a Nota Fiscal para o Hospital. A Nota Fiscal é certificada pelo gestor (titular ou substituto) e, posteriormente, encaminhada para a Seção de Execução Financeira.",
                                        8: "Liquidação da Nota Fiscal no Sistema Integrado de Administração Financeira do Governo Federal (SIAFI) 🖥️, conforme o estabelecido no artigo 63 da Lei 4.320/64.",
                                        9: "Pagamento autorizado pelo Ordenador de Despesas. Nessa etapa, o pagamento demora, em média, um dia útil após a aprovação para ser creditado em conta-corrente. Outrossim, salienta-se que, por ocasião dos pagamentos, são realizados os abatimentos tributários devidos."
                                    }

                                    # 4. Mapeamento de meses para siglas
                                    meses_map = {
                                        1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR", 5: "MAI", 6: "JUN",
                                        7: "JUL", 8: "AGO", 9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ"
                                    }

                                    # 5. Construção do Corpo do Resumo
                                    resumo_corpo = ""
                                    etapas_ativas = sorted(df_proc_fisc['status'].unique())
                                    
                                    for st_id in etapas_ativas:
                                        if st_id >= 9: continue # Ignora processos já pagos
                                        
                                        df_etapa = df_proc_fisc[df_proc_fisc['status'] == st_id].copy()
                                        
                                        if not df_etapa.empty:
                                            nome_da_etapa = mapa_status_nomes.get(st_id, f"Etapa {st_id}")
                                            total_etapa = df_etapa['v_liq'].sum()
                                            descricao = explica_etapa.get(st_id, "")
                                            
                                            resumo_corpo += f"\n🔹 **{nome_da_etapa.upper()}**\n"
                                            resumo_corpo += f"   - {descricao}\n"
                                            resumo_corpo += f"   - Volume Total na Etapa: R$ {total_etapa:,.2f}\n"
                                            
                                            # Agrupamento e formatação por Mês de Entrada (Ex: JAN2026)
                                            # Agora a dt_ordem existe e o sorted funcionará perfeitamente!
                                            competencias_ordenadas = sorted(df_etapa['dt_ordem'].dropna().unique())
                                            
                                            for dt in competencias_ordenadas:
                                                df_comp = df_etapa[df_etapa['dt_ordem'] == dt]
                                                
                                                # Converte data em JAN2026
                                                mes_sigla = meses_map.get(dt.month, "INV")
                                                ano_ref = dt.year
                                                rotulo_entrada = f"{mes_sigla}{ano_ref}"
                                                
                                                lista_faturas = ", ".join(df_comp['Numero_da_fatura'].astype(str).tolist())
                                                subtotal_comp = df_comp['v_liq'].sum()
                                                
                                                resumo_corpo += f"     ➔ Mês de entrada no HNBra: {rotulo_entrada}: Faturas [{lista_faturas}] — Subtotal: R$ {subtotal_comp:,.2f}\n"
                                            
                                            resumo_corpo += "───────────────────────────────────────\n"

                                    # 6. Template do E-mail (Mantendo seu texto oficial)
                                    msg_final = f"""Prezado (a) representante da empresa {ose_sel},

Cumprimentando-o (a) cordialmente, seguem abaixo algumas orientações, bem como um panorama atualizado de seus processos ora em trâmite no Hospital Naval de Brasília (HNBra):

📌 (i) DO CRITÉRIO DE PAGAMENTO
Este hospital realiza a emissão das Notas de Empenho em estrita ordem cronológica, mediante disponibilidade orçamentária, a partir da data da entrada da(s) fatura(s) em nossa Secretaria. Nesse contexto, é útil uma análise literal da Lei 14.133/21 (Art. 141): 
                        
"No dever de pagamento pela Administração, será observada a ordem cronológica para cada fonte diferenciada de recursos [...]"

📊 (ii) COMPOSIÇÃO ATUAL DOS PROCESSOS
{resumo_corpo}

🚀 (iii) PRÓXIMOS PASSOS
As faturas supracitadas seguem em fluxo contínuo de processamento. Assim que concluídas as etapas de conferência e reserva orçamentária, as Notas de Empenho correspondentes a cada fatura serão encaminhadas para a respectiva emissão de Nota Fiscal e posterior liquidação.

Reitera-se nosso compromisso com a transparência e eficiência em nossos atos administrativos. 

Estamos à disposição para eventuais esclarecimentos. Gratos pela distinta parceria! 🤝⚓🇧🇷 

Cordialmente,
"""
                                    # Salva na sessão para persistência
                                    st.session_state['panorama_gerado'] = msg_final

                        # 7. Exibição do Resultado
                        if 'panorama_gerado' in st.session_state:
                            st.success("""✅ **Panorama Gerencial pronto para cópia!** 🫡🇧🇷""")
                            
                            st.code(st.session_state['panorama_gerado'], language="text")



                    else:
                        st.info(f"Nenhum processo encontrado para {ose_sel}.")

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
                    # --- 1. GATILHO DE ATUALIZAÇÃO (O SEGREDO) ---
                    if "ne_anterior" not in st.session_state:
                        st.session_state.ne_anterior = ne_alvo
                    
                    if st.session_state.ne_anterior != ne_alvo:
                        # Identificamos as chaves da NE que está saindo de cena
                        chave_sub_velha = f"sub_{st.session_state.ne_anterior}"
                        chave_body_velha = f"body_{st.session_state.ne_anterior}"
                        
                        # APAGAMOS a memória delas para não contaminar a próxima
                        if chave_sub_velha in st.session_state:
                            del st.session_state[chave_sub_velha]
                        if chave_body_velha in st.session_state:
                            del st.session_state[chave_body_velha]
                        
                        # Atualizamos o rastreador e damos o "tranco" no motor
                        st.session_state.ne_anterior = ne_alvo
                        st.rerun() 

                    # --- 2. CÁLCULO DOS DADOS (UMA ÚNICA VEZ) ---
                    df_ne_fisc = df_s6[df_s6['ne'] == ne_alvo].copy()
                    ose_txt = df_ne_fisc['ose'].iloc[0]
                    v_total = df_ne_fisc['valor_liquido'].apply(limpar_valor).sum()
                    faturas_txt = ", ".join(df_ne_fisc['Numero_da_fatura'].astype(str).unique())
                    cnpj_alvo = str(df_ne_fisc['cnpj'].iloc[0]).split('.')[0].zfill(14)
                    
                    # Busca contatos na Tabela-A
                    # --- BUSCA DE CONTATOS (BLINDAGEM TÁTICA) ---
                    # --- BUSCA DE CONTATOS (AGORA COM O NOME CORRETO) ---
                    try:
                        # 1. Carrega a aba
                        df_tabela_a = pd.DataFrame(sh.worksheet(ABA_TABELA_A).get_all_records())
                        df_tabela_a.columns = df_tabela_a.columns.str.strip()
                        
                        # 🛡️ Função para garantir 14 dígitos (Zeros à Esquerda)
                        def normalizar_cnpj(valor):
                            apenas_numeros = "".join(filter(str.isdigit, str(valor)))
                            return apenas_numeros.zfill(14) if apenas_numeros else ""

                        # 2. Normalização
                        cnpj_busca = normalizar_cnpj(cnpj_alvo)
                        # AQUI ESTAVA O ERRO: Usei o nome completo df_tabela_a em todo lugar
                        df_tabela_a['CNPJ_LIMPO'] = df_tabela_a['CNPJ'].apply(normalizar_cnpj)
                        
                        # 3. Busca Exata
                        linha_ose = df_tabela_a[df_tabela_a['CNPJ_LIMPO'] == cnpj_busca]
                        
                        if not linha_ose.empty:
                            email_destino = str(linha_ose.iloc[0].get('E-mail Principal da OSE', "faturamento@ose.com")).strip()
                            email_titular = str(linha_ose.iloc[0].get('E-mail do Gestor Titular', "")).strip()
                            email_substituto = str(linha_ose.iloc[0].get('E-mail do Gestor Substituto', "")).strip()
                            
                            # Dica: Se quiser conferir se achou, descomente a linha abaixo temporariamente:
                            # st.write(f"DEBUG: Encontrado {email_destino}")
                        else:
                            email_destino = "faturamento@ose.com"
                            email_titular, email_substituto = "", ""

                        email_exec = "hnbra.execucaofinanceira@gmail.com"
                        
                    except Exception as e:
                        # Se der erro, mostra o erro para sabermos o que é
                        st.error(f"Erro na Tabela-A: {e}")
                        email_destino = "faturamento@ose.com"
                        email_titular, email_substituto = "", ""
                        email_exec = "hnbra.execucaofinanceira@gmail.com"

                    st.markdown(f"#### 📝 Gestão da NE: **{ne_alvo}** ({ose_txt})")
                    
                    # --- INÍCIO DAS COLUNAS (Alinhamento corrigido) ---
                    col_f1, col_f2 = st.columns(2)
                    
                    with col_f1:
                        st.markdown("##### 📤 1. Informar Nota Fiscal")
                        nf_in = st.text_input("Número da NF recebida:", placeholder="Ex: 550/2026", key=f"nf_input_{ne_alvo}")
                        
                        if st.button("💾 Registrar NF no SISAFA NAVAL", use_container_width=True, key=f"btn_nf_{ne_alvo}"):
                            if nf_in:
                                with st.spinner("Gravando nota..."):
                                    # --- CORREÇÃO AQUI: Definindo a aba de processos antes de usar ---
                                    aba_p = sh.worksheet(ABA_PROCESSOS) 
                                    
                                    # O loop agora saberá quem é 'aba_p'
                                    for nup_item in df_ne_fisc['nup'].tolist():
                                        cell = aba_p.find(nup_item)
                                        if cell:
                                            # Coluna 16 (ajuste se a coluna da NF for outra na sua planilha)
                                            aba_p.update_cell(cell.row, 16, nf_in) 
                                            registrar_acao(nup_item, "N/A", "NF_INFORMADA", f"NF: {nf_in}")
                                
                                st.success(f"✅ NF {nf_in} registrada com sucesso!")
                                import time # Garante que o time está disponível para o sleep
                                time.sleep(1)
                                st.rerun()
                            else:
                                st.warning("⚠️ Informe o número da NF.")

                    with col_f2:
                        st.markdown("#### 📧 2. Solicitação de Nota Fiscal")

                        # --- BOTÃO DE SINCRONIZAÇÃO ---
                        if st.button("🔄 Sincronizar Texto", help="Força a atualização dos dados desta NE", key=f"sync_{ne_alvo}", use_container_width=True):
                            for k in [f"sub_{ne_alvo}", f"body_{ne_alvo}"]:
                                if k in st.session_state:
                                    del st.session_state[k]
                            st.rerun()

                        # --- 1. LÓGICA DE CÓPIAS (CC) - BUSCA DINÂMICA POR NIP ---
                        # Pegamos o NIP de quem está logado
                        nip_logado = str(st.session_state.get('user_id', '')).strip()
                        
                        # Se o e-mail não estiver na memória, buscamos na base de usuários
                        if 'email_usuario_logado' not in st.session_state or not st.session_state['email_usuario_logado']:
                            try:
                                df_u = pd.DataFrame(sh.worksheet("SISAFA-NAVAL-Usuarios").get_all_records())
                                # Identifica se a coluna chama 'E-mail' ou 'Email'
                                col_mail_u = 'E-mail' if 'E-mail' in df_u.columns else 'Email'
                                match_u = df_u[df_u['NIP'].astype(str).str.strip() == nip_logado]
                                
                                if not match_u.empty:
                                    st.session_state['email_usuario_logado'] = match_u[col_mail_u].values[0]
                                    if 'NOME' in match_u.columns:
                                        st.session_state['nome_usuario'] = match_u['NOME'].values[0]
                                else:
                                    st.session_state['email_usuario_logado'] = ''
                            except Exception:
                                st.session_state['email_usuario_logado'] = ''

                        usuario_atual = st.session_state.get('email_usuario_logado', '')
                        
                        # Criamos a lista com o seu e-mail + os titulares recuperados no topo do script
                        lista_cc_bruta = [usuario_atual, email_titular, email_substituto, email_exec]
                        
                        # Limpeza de 'nan', vazios e sujeiras
                        lista_cc_limpa = [str(u).strip().rstrip(',') for u in lista_cc_bruta if u and str(u).lower().strip() != 'nan']
                        
                        # Junta tudo para os cabeçalhos do e-mail
                        cc_string = ", ".join(lista_cc_limpa).strip().rstrip(',')
                        email_destino_limpo = str(email_destino).strip().rstrip(',')

                        # --- 2. MONTAGEM DO TEXTO (Mantém seu texto padrão) ---
                        assunto_fresco = f"Solicitação de Nota Fiscal para pagamento – Hospital Naval de Brasília"
                        

                        # Busca o nome real uma única vez para não pesar o sistema
                        if 'nome_usuario' not in st.session_state:
                            df_u = pd.DataFrame(sh.worksheet("SISAFA-NAVAL-Usuarios").get_all_records())
                            st.session_state['nome_usuario'] = next(iter(df_u[df_u['E-mail'] == usuario_atual]['NOME']), 'Fiscal de Contratos')

                        corpo_fresco = (
                            f"🛑FAVOR NÃO RESPONDER DIRETAMENTE AO E-MAIL 'hnbra.execucaofinanceira@gmail.com' SEM DEIXAR OS ENDEREÇADOS ELENCADOS NA CÓPIA🛑\n\n"
                            f"🛑ISSO ACARRETA ATRASOS NO PAGAMENTO! FAVOR SELECIONAR A OPÇÃO 'RESPONDER A TODOS'!🛑\n\n"
                            f"Prezados do (a) {ose_txt}, CNPJ {cnpj_alvo}, esperamos que este e-mail os encontre bem.\n\n"
                            f"O HNBra solicita, por gentileza, a emissão de Nota Fiscal referente aos serviços prestados por essa Organização de Saúde Extra – Marinha do Brasil, conforme acordo vigente.\n\n"
                            f"💲 Valor total: R$ {v_total:,.2f}\n"
                            f"🖨️ Fatura(s) / protocolo(s) / remessa(s): {faturas_txt} de {datetime.now().year}.\n"
                            f"📃 Nota de empenho: {datetime.now().year} NE {str(ne_alvo)[-6:] if ne_alvo else ''}.\n\n"
                            f"⚠️ ATENÇÃO! ⚠️ Não emitir a Nota Fiscal, sem antes, conferir os valores e números das faturas, valor total da remessa, CNPJ correto e relatórios de glosa quando houver! A emissão incorreta de NF acarreta atrasos ao pagamento e problemas administrativos junto ao fisco.\n\n"
                            f"📃📃 Dados para Nota Fiscal 📃📃\n"
                            f"Razão Social: Hospital Naval de Brasília\n"
                            f"Endereço: SEPS Q 711/911 - Asa Sul, Brasília - DF, 70390-115\n"
                            f"Telefone de Contato: (61) 3445-7303\n"
                            f"CNPJ: 00.394.502/0060-02\n\n"
                            f"👉 Obs.: Participo que não confirmo o recebimento de todas as notas fiscais por causa da grande demanda aqui e para não atrasar ainda mais os pagamentos, mas quando não as recebo, solicito novamente.\n\n"
                            f"Sempre priorizo os pagamentos, quando chega dotação orçamentária. Assim, somente respondo os emails, em tempo, com notas fiscais de valores 'errados' ou de casos extremamente urgentes, e no final da remessa, verifico todas as demais pendências.\n\n"
                            f"Peço a gentileza de aguardar!\n\n"
                            f"Por favor, anexar a Nota Fiscal na mensagem de solicitação por e-mail.\n\n"
                            f"Fiscalização de contratos\n"
                            f"⚓ Hospital Naval de Brasília ⚓"
                        )

                        # --- 3. INTERFACE ---
                        with st.container(border=True):
                            st.markdown(f"**Para:** {email_destino_limpo}")
                            st.markdown(f"**CC:** {cc_string if cc_string else '---'}")
                            
                            st.divider() 
                            
                            assunto_final = st.text_input("Assunto:", value=assunto_fresco, key=f"sub_{ne_alvo}")
                            msg_final = st.text_area("Corpo da mensagem:", value=corpo_fresco, height=450, key=f"body_{ne_alvo}")

                            if st.button("📧 Disparar Solicitação Oficial", use_container_width=True, key=f"btn_mail_{ne_alvo}"):
                                with st.spinner("🚀 Enviando e-mail oficial..."):
                                    # --- CHAMADA DA FUNÇÃO REAL ---
                                    sucesso = enviar_email_generico(
                                        destinatario=email_destino_limpo,
                                        assunto=assunto_final,
                                        corpo=msg_final,
                                        cc=cc_string
                                    )
                                    
                                    if sucesso:
                                        st.success(f"✅ E-mail enviado com sucesso para {email_destino_limpo}!")
                                        with st.spinner("📑 Atualizando registros de auditoria..."):
                                            # Percorre todos os processos (NUPs) vinculados a esta NE
                                            for nup_item in df_ne_fisc['nup'].tolist():
                                                try:
                                                    # Recupera os dados do processo para o log
                                                    dados_nup = df[df['nup'] == nup_item].iloc[0]
                                                    fatura_n = dados_nup['Numero_da_fatura']
                                                    # Garante que pegamos o valor líquido formatado
                                                    valor_momento = dados_nup.get('v_liq', dados_nup.get('valor_liquido', 0))
                                                    status_atual = str(dados_nup['status']).replace('.0', '')

                                                    # 1. REGISTRA A AÇÃO (Log de Eventos)
                                                    registrar_acao(
                                                        nup_item, 
                                                        fatura_n, 
                                                        "SOLICITACAO_NF_ENVIADA", 
                                                        f"E-mail enviado para: {email_destino_limpo}"
                                                    )

                                                    # 2. REGISTRA O HISTÓRICO (Cronômetro de Produtividade)
                                                    # Mantemos o status de origem e destino iguais (ex: 6 -> 6) 
                                                    # apenas para marcar que houve interação no processo hoje.
                                                    registrar_historico(
                                                        nup_item,
                                                        fatura_n,
                                                        status_atual,
                                                        status_atual,
                                                        valor_momento,
                                                        f"Solicitação oficial de NF enviada para {email_destino_limpo}"
                                                    )
                                                except Exception as e:
                                                    st.error(f"Erro ao registrar log para o NUP {nup_item}: {e}")
                                        
                                        st.info("💡 Registros de auditoria atualizados com sucesso.")
                                        time.sleep(1.5)
                                        st.rerun()

                            # --- BOTÃO DE DEVOLUÇÃO (Estorno de NE) ---
                            st.divider()
                            with st.expander("🚨 Detectou erro na NE? (Devolver p/ Execução)"):
                                st.write("""
                                ⚠️ **Orientações** ⚠️

                                Prezado (a), caso tenha constatado algum erro na Nota de Empenho, tais como: 
                                
                                🛑 Valor incorreto;

                                🛑 Não sou fiscal (titular ou substituto) da referida Nota de Empenho;
                                
                                🛑 CNPJ incorreto;
                                
                                🛑 A descrição do empenho está inadequada, entre outros, clique em **Confirmar Devolução p/ Execução** 🔘👈.  
                                
                                Quaisquer dúvidas, entre em contrato com a Seção de Execução Financeira! Estamos a disposição para quaisquer dúvidas. ⚓🇧🇷🫡
                                
                                📞 8916-7349 (Retelma)
                                
                                📞 8916-7345 (Retelma)
                                
                                📞 8916-7361 (Retelma) 
                                """)

                                motivo_estorno = st.text_area("Motivo da devolução:", placeholder="Ex: Valor da NE não confere com o líquido da auditoria.", key=f"motivo_dev_{ne_alvo}")
                                
                                if st.button("↩️ Confirmar Devolução p/ Execução", type="primary", use_container_width=True, key=f"btn_dev_{ne_alvo}"):
                                    if motivo_estorno:
                                        with st.spinner("Estornando processos..."):
                                            try:
                                                aba_p = sh.worksheet("SISAFA-NAVAL-processos")
                                                # Captura o NIP do usuário logado (assumindo que está no session_state)
                                                usuario_atual = st.session_state.get('usuario_nip', 'N/A')
                                                
                                                for nup_item in df_ne_fisc['nup'].tolist():
                                                    nup_str = str(nup_item).strip()
                                                    cell = aba_p.find(nup_str)
                                                    
                                                    if cell:
                                                        # --- CAPTURA DE DADOS PARA O HISTÓRICO ---
                                                        # Lemos a linha inteira para pegar status antigo e valor
                                                        dados_linha = aba_p.row_values(cell.row)
                                                        
                                                        # Ajuste os índices [X] abaixo de acordo com sua planilha real:
                                                        fatura_n = dados_linha[4]  # Coluna E (Numero_da_fatura)
                                                        status_atual = dados_linha[10] # Coluna K (status)
                                                        valor_momento = dados_linha[7] # Coluna H (valor_no_momento/líquido)
                                                        
                                                        # 1. Atualiza na aba de PROCESSOS (Retorna para Status 5)
                                                        aba_p.update_cell(cell.row, 11, 5) 
                                                        
                                                        # 2. Registra na aba de AÇÕES (Log de auditoria)
                                                        registrar_acao(nup_str, fatura_n, "NE_DEVOLVIDA", f"Motivo: {motivo_estorno}")
                                                        
                                                        # 3. REGISTRA NA ABA HISTÓRICO (O que você pediu)
                                                        # A função registrar_historico deve seguir a ordem: 
                                                        # nup, fatura, orig, dest, valor, obs
                                                        registrar_historico(
                                                            nup_str, 
                                                            fatura_n, 
                                                            status_atual, 
                                                            "5", 
                                                            valor_momento, 
                                                            f"ESTORNO: {motivo_estorno}"
                                                        )
                                                
                                                st.success(f"✅ NE {ne_alvo} devolvida com sucesso!")
                                                time.sleep(1.5)
                                                st.rerun()
                                                
                                            except Exception as e:
                                                st.error(f"Erro no estorno: {e}")
                                    else:
                                        st.error("⚠️ É obrigatório informar o motivo para a devolução.")    


        # --- 3. ABA: RELACIONAMENTO (Módulo Fiscalização) ---
        with tab_rel:
            st.subheader("💬 Central de Relacionamento (Gestão de OSEs)")
            
            try:
                # 1. Carregamos as mensagens da planilha
                aba_msg = sh.worksheet(ABA_MENSAGENS)
                dados_brutos = aba_msg.get_all_records()
                df_msg = pd.DataFrame(dados_brutos)
                
                if df_msg.empty:
                    st.info("Nenhuma mensagem registrada no sistema.")
                else:
                    # --- VACINA NAS MENSAGENS ---
                    df_msg['remetente_limpo'] = df_msg['remetente'].astype(str).str.split('.').str[0].str.strip().str.zfill(14)

                    # --- FILTRO DE SEGURANÇA ---
                    if not is_global:
                        # PROTEÇÃO: Se por algum motivo o df_fiscal perdeu a coluna, criamos aqui agora
                        if 'CNPJ_LIMPO' not in df_fiscal.columns:
                            # Tentamos achar a coluna original do CNPJ (pode ser 'CNPJ' ou 'CNPJ_OSE' dependendo da aba)
                            col_base = 'CNPJ' if 'CNPJ' in df_fiscal.columns else df_fiscal.columns[0]
                            df_fiscal['CNPJ_LIMPO'] = df_fiscal[col_base].astype(str).str.split('.').str[0].str.strip().str.zfill(14)
                        
                        cnpjs_meus = df_fiscal['CNPJ_LIMPO'].tolist()
                        df_msg_filtrado = df_msg[df_msg['remetente_limpo'].isin(cnpjs_meus)].copy()
                    else:
                        df_msg_filtrado = df_msg.copy()

                    if df_msg_filtrado.empty:
                        st.info("📭 Nenhuma mensagem pendente das suas OSEs.")
                    else:
                        # 2. Métricas (Ajustadas para o nome correto 'status_msg')
                        pendentes = len(df_msg_filtrado[df_msg_filtrado['status_msg'] == 'PENDENTE'])
                        c1, c2 = st.columns(2)
                        c1.metric("Total de Mensagens", len(df_msg_filtrado))
                        c2.metric("📩 Pendentes", pendentes, delta_color="inverse")

                        st.divider()

                        # 3. Tabela de Mensagens
                        st.write("**📥 Histórico de Interações:**")
                        cols_vistas = ['Numero_da_fatura', 'nup', 'remetente', 'setor_destino', 'status_msg']
                        st.dataframe(
                            df_msg_filtrado[cols_vistas].sort_values(by='status_msg', ascending=False),
                            use_container_width=True,
                            hide_index=True
                        )

                        st.markdown("---")

                        # 4. Área de Resposta
                        st.markdown("### ✍️ Responder ou Intervir")
                        df_msg_filtrado['label_selecao'] = (
                            "Fatura: " + df_msg_filtrado['Numero_da_fatura'].astype(str) + 
                            " | ID: " + df_msg_filtrado['id_mensagem'].astype(str)
                        )
                        
                        selecao_msg = st.selectbox("Selecione a mensagem para responder:", [""] + df_msg_filtrado['label_selecao'].tolist(), key="sb_rel_fisc_definitivo")
                        
                        if selecao_msg:
                            dados_m = df_msg_filtrado[df_msg_filtrado['label_selecao'] == selecao_msg].iloc[0]
                            id_msg_alvo = str(dados_m['id_mensagem'])
                            
                            with st.container(border=True):
                                st.write(f"🏢 **OSE:** {dados_m['remetente']}")
                                st.chat_message("user").write(f"**Dúvida:** {dados_m['texto']}")
                                
                                resp_fisc = st.text_area("Sua Resposta Oficial (Fiscal):", key="txt_fisc_final")
                                
                                if st.button("📤 ENVIAR RESPOSTA DO FISCAL", use_container_width=True):
                                    if resp_fisc:
                                        with st.spinner("Enviando..."):
                                            registrar_acao(dados_m['nup'], dados_m['Numero_da_fatura'], "RESPOSTA_FISCAL", f"Fiscal respondeu ID {id_msg_alvo}")
                                            
                                            # Busca a linha e atualiza as 10 colunas
                                            celula = aba_msg.find(id_msg_alvo)
                                            linha_idx = celula.row
                                            agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                            
                                            aba_msg.update_cell(linha_idx, 8, agora)         # data_resposta
                                            aba_msg.update_cell(linha_idx, 9, "RESPONDIDO")    # status_msg
                                            aba_msg.update_cell(linha_idx, 10, str(st.session_state.user_id)) # NIP
                                            
                                            st.success("Resposta enviada!")
                                            time.sleep(1.5)
                                            st.rerun()
            except Exception as e:
                st.error(f"Erro no módulo de relacionamento: {e}")
    





    elif st.session_state.modulo_ativo == "GERENCIAL":
        st.header("📈 Análise Estratégica")

        # --- 1. DEFINIÇÕES GLOBAIS ---
        mapa_meses_abrev = {1:'JAN', 2:'FEV', 3:'MAR', 4:'ABR', 5:'MAI', 6:'JUN', 
                            7:'JUL', 8:'AGO', 9:'SET', 10:'OUT', 11:'NOV', 12:'DEZ'}

        cats_oficiais = [
            "OSE", 
            "HOSPITAL DAS FORÇAS ARMADAS (HFA)", 
            "Base Administrativa do Comando de Operações Especiais (160098)", 
            "Base Aérea de Anápolis (120624)", 
            "HFAB (120096)"
        ]

        # --- FUNÇÃO MESTRE UNIFICADA ---
        def gerar_tabela_gerencial(df_input, lista_status):
            df_temp = df_input.copy()
            # Vacina de tipos para evitar erros de processamento
            df_temp['status'] = pd.to_numeric(df_temp['status'], errors='coerce').fillna(0)
            df_temp['mes_competencia'] = pd.to_numeric(df_temp['mes_competencia'], errors='coerce').fillna(0).astype(int)
            df_temp['ano_competencia'] = pd.to_numeric(df_temp['ano_competencia'], errors='coerce').fillna(0).astype(int)
            
            # Filtro pelos status desejados (ex: [4] ou [1, 2, 3])
            df_c = df_temp[df_temp['status'].isin(lista_status)].copy()
            
            if df_c.empty:
                return pd.DataFrame(), pd.DataFrame()

            df_c['v_liq'] = df_c['valor_liquido'].apply(limpar_valor)
            
            def categorizar(nome):
                n = str(nome).upper()
                if "HOSPITAL DAS FORÇAS ARMADAS" in n or "HFA" in n: return cats_oficiais[1]
                elif "160098" in n or "OPERAÇÕES ESPECIAIS" in n: return cats_oficiais[2]
                elif "120624" in n or "BASE AÉREA DE ANÁPOLIS" in n: return cats_oficiais[3]
                elif "120096" in n or "HFAB" in n: return cats_oficiais[4]
                return cats_oficiais[0]

            df_c['Categoria'] = df_c['ose'].apply(categorizar)
            df_long = df_c.groupby(['Categoria', 'ano_competencia', 'mes_competencia'])['v_liq'].sum().reset_index()
            
            # Pivotagem e Reindexação
            df_pivot = df_long.pivot(index='Categoria', columns=['ano_competencia', 'mes_competencia'], values='v_liq').fillna(0.0)
            df_pivot = df_pivot.sort_index(axis=1, level=[0, 1])
            df_pivot = df_pivot.reindex(cats_oficiais).fillna(0.0)
            
            # Totais
            total_linha = df_pivot.sum(axis=1)
            novas_labels = [(int(ano), mapa_meses_abrev.get(int(mes), "???")) for ano, mes in df_pivot.columns]
            df_pivot.columns = pd.MultiIndex.from_tuples(novas_labels, names=['Ano', 'Mês'])
            df_pivot = df_pivot.loc[:, (df_pivot != 0).any(axis=0)]
            df_pivot.insert(0, ('TOTAL', 'ACUMULADO'), total_linha)
            df_pivot.loc['TOTAL GERAL'] = df_pivot.sum()
            
            return df_pivot, df_long

        tab_fin, tab_prod, tab_est = st.tabs(["💰 Situação Financeira", "⏱️ Produtividade", "📂 Estrutura"])

        with tab_fin:
            # === SEÇÃO 1: STATUS 4 (AUDITADOS) ===
            st.subheader("📌 1. Créditos orçamentários comprometidos auditados")
            df_creditos, df_grafico_4 = gerar_tabela_gerencial(df, [4])

            if df_creditos.empty:
                st.info("Nenhuma fatura aguardando emissão de NE para exibir.")
            else:
                st.dataframe(df_creditos.style.format("R$ {:,.2f}").set_table_styles([
                    {'selector': 'th', 'props': [('background-color', '#2e6b54'), ('color', 'white'), ('font-weight', 'bold')]}
                ]), use_container_width=True)

                # Histograma Status 4
                df_grafico_4['Competência'] = df_grafico_4.apply(lambda x: f"{mapa_meses_abrev[int(x['mes_competencia'])]}/{str(int(x['ano_competencia']))[2:]}", axis=1)
                df_grafico_4['sort_key'] = df_grafico_4['ano_competencia'] * 100 + df_grafico_4['mes_competencia']
                df_grafico_4 = df_grafico_4.sort_values('sort_key')
                
                fig4 = px.bar(df_grafico_4, x='Competência', y='v_liq', color='Categoria', title="Dívida Auditada",
                              color_discrete_map={cats_oficiais[0]: "#2e6b54", cats_oficiais[1]: "#cba30c"}, barmode='stack')
                st.plotly_chart(fig4, use_container_width=True)

            st.divider()

            # === SEÇÃO 2: STATUS 1, 2, 3 (EM FLUXO) ===
            st.subheader("⏳ 2. Evolução dos valores comprometidos em auditagem")
            df_pendentes, df_grafico_pend = gerar_tabela_gerencial(df, [1, 2, 3])

            if df_pendentes.empty:
                st.info("Não há faturas em processo de auditagem no momento.")
            else:
                st.dataframe(df_pendentes.style.format("R$ {:,.2f}").set_table_styles([
                    {'selector': 'th', 'props': [('background-color', '#1e3d33'), ('color', 'white'), ('font-weight', 'bold')]}
                ]), use_container_width=True)
                
                # Histograma Status 1, 2, 3
                df_grafico_pend['Competência'] = df_grafico_pend.apply(lambda x: f"{mapa_meses_abrev[int(x['mes_competencia'])]}/{str(int(x['ano_competencia']))[2:]}", axis=1)
                df_grafico_pend['sort_key'] = df_grafico_pend['ano_competencia'] * 100 + df_grafico_pend['mes_competencia']
                df_grafico_pend = df_grafico_pend.sort_values('sort_key')

                fig_pend = px.bar(df_grafico_pend, x='Competência', y='v_liq', color='Categoria', title="Volume em Auditagem",
                                  color_discrete_sequence=px.colors.qualitative.Pastel, barmode='stack')
                st.plotly_chart(fig_pend, use_container_width=True)

            # === SEÇÃO 3: PANORAMA GERAL (PIZZA) ===
            st.divider()
            st.subheader("📊 3. Panorama Geral das faturas 📑📑")

            # 1. Dicionário tradutor de Status
            mapa_status_nomes = {                
                1: "1. 📥 Cadastrada (SECOM)",
                2: "2. 🩺 Em Auditagem",
                3: "3. ✅ Auditada",
                4: "4. 💰 Aguardando NE",
                5: "5. 🏦 Empenhada",
                6: "6. 📝 Aguardando NF",
                7: "7. ⏳ Em liquidação",
                8: "8. 🖥️ Liquidada",
                9: "9. 💸 Paga "
            }

            # --- PALETA DE CORES TÁTICA E HARMONIOSA ---
            # Cada fase ganha uma cor fixa. Tons que combinam, mas não se confundem.
            cores_fases = {
                "1. 📥 Cadastrada (SECOM)": "#3498db",  # Azul Claro
                "2. 🩺 Em Auditagem": "#e67e22",      # Laranja
                "3. ✅ Auditada": "#2ecc71",          # Verde Claro
                "4. 💰 Aguardando NE": "#e74c3c",       # Vermelho
                "5. 🏦 Empenhada": "#9b59b6",         # Roxo
                "6. 📝 Aguardando NF": "#f1c40f",       # Amarelo/Dourado
                "7. ⏳ Em liquidação": "#1abc9c",       # Turquesa
                "8. 🖥️ Liquidada": "#34495e",         # Azul Marinho/Cinza
                "9. 💸 Paga ": "#27ae60",             # Verde Escuro
                "Outros / Desconhecido": "#95a5a6"      # Cinza Claro
            }

            # 2. Preparação dos Dados
            df_pano = df.copy()
            df_pano['status_num'] = pd.to_numeric(df_pano['status'], errors='coerce').fillna(-1).astype(int)
            df_pano['Fase Atual'] = df_pano['status_num'].map(mapa_status_nomes).fillna('Outros / Desconhecido')
            df_pano['v_liq'] = df_pano['valor_liquido'].apply(limpar_valor)

            # 3. Agrupamento de Dados (Conta faturas e soma valores)
            df_resumo_pano = df_pano.groupby('Fase Atual').agg(
                Qtd_Faturas=('nup', 'count'),  
                Volume_RS=('v_liq', 'sum')     
            ).reset_index()

            # Remove status vazios e ORDENA o dataframe (Para a legenda ficar 1, 2, 3...)
            df_resumo_pano = df_resumo_pano[df_resumo_pano['Qtd_Faturas'] > 0].sort_values('Fase Atual')

            if df_resumo_pano.empty:
                st.info("Não há dados suficientes para gerar o panorama.")
            else:
                # 4. Construção dos Gráficos Lado a Lado
                col_g1, col_g2 = st.columns(2)

                # Ordem forçada para a legenda
                ordem_legenda = sorted(df_resumo_pano['Fase Atual'].tolist())

                with col_g1:
                    fig_qtd = px.pie(
                        df_resumo_pano, 
                        values='Qtd_Faturas', 
                        names='Fase Atual', 
                        hole=0.4, 
                        title="Distribuição por Quantidade (Nº de Faturas)",
                        color='Fase Atual',           # Diz ao Plotly para colorir baseado no nome
                        color_discrete_map=cores_fases, # Aplica o nosso dicionário de cores
                        category_orders={"Fase Atual": ordem_legenda} # Força a ordem 1, 2, 3...
                    )
                    fig_qtd.update_traces(textposition='inside', textinfo='percent+value')
                    # Se quiser esconder a legenda do primeiro gráfico para poupar espaço:
                    fig_qtd.update_layout(showlegend=False) 
                    st.plotly_chart(fig_qtd, use_container_width=True)

                with col_g2:
                    fig_vol = px.pie(
                        df_resumo_pano, 
                        values='Volume_RS', 
                        names='Fase Atual', 
                        hole=0.4,
                        title="Distribuição por Volume Financeiro (R$)",
                        color='Fase Atual',
                        color_discrete_map=cores_fases,
                        category_orders={"Fase Atual": ordem_legenda}
                    )
                    fig_vol.update_traces(
                        textposition='inside', 
                        textinfo='percent',
                        hovertemplate="<b>%{label}</b><br>Volume: R$ %{value:,.2f}<br>Representação: %{percent}"
                    )
                    # Mantém a legenda apenas no segundo gráfico, já ordenada
                    st.plotly_chart(fig_vol, use_container_width=True)

                # 5. Tabela Resumo Expandível
                with st.expander("Ver dados detalhados do Panorama"):
                    df_resumo_pano_show = df_resumo_pano.copy()
                    df_resumo_pano_show['Volume (R$)'] = df_resumo_pano_show['Volume_RS'].apply(lambda x: f"R$ {x:,.2f}")
                    df_resumo_pano_show = df_resumo_pano_show.drop(columns=['Volume_RS'])
                    
                    st.dataframe(df_resumo_pano_show, use_container_width=True, hide_index=True)

            
            # =======================================================
            # === FUNÇÃO AUXILIAR TÁTICA (Agrupa as pequenas fatias) ===
            # =======================================================
            def agrupar_top_n(df_agrupado, col_nome, col_valor, top_n=10):
                df_agrupado = df_agrupado.sort_values(col_valor, ascending=False)
                if len(df_agrupado) > top_n:
                    top_df = df_agrupado.iloc[:top_n].copy()
                    outros_valor = df_agrupado.iloc[top_n:][col_valor].sum()
                    outros_df = pd.DataFrame({col_nome: ['OUTRAS EMPRESAS'], col_valor: [outros_valor]})
                    return pd.concat([top_df, outros_df], ignore_index=True)
                return df_agrupado

            # =======================================================
            # === SEÇÃO 4: PANORAMA FINANCEIRO POR EMPRESA (GERAL) ===
            # =======================================================
            st.divider()
            st.subheader("🏢 4. Panorama Financeiro por Empresa")

            df_sec4 = df.copy()
            if 'v_liq' not in df_sec4.columns:
                df_sec4['v_liq'] = df_sec4['valor_liquido'].apply(limpar_valor)

            df_empresas_geral = df_sec4.groupby('ose')['v_liq'].sum().reset_index()
            df_empresas_geral = df_empresas_geral[df_empresas_geral['v_liq'] > 0]

            if df_empresas_geral.empty:
                st.info("Não há dados financeiros vinculados às empresas no momento.")
            else:
                # Aplica a tática de agrupar o "Top 10" e esconder as menores em "Outras"
                df_pizza_geral = agrupar_top_n(df_empresas_geral, 'ose', 'v_liq', top_n=10)

                # Gráfico limpo e gigante (sem linhas, porcentagem dentro)
                fig_empresas = px.pie(
                    df_pizza_geral, 
                    values='v_liq', 
                    names='ose', 
                    hole=0.45, # Furo maior, estilo Donut moderno
                    title="As 10 Maiores Empresas em Volume Financeiro Total",
                    color_discrete_sequence=px.colors.qualitative.Prism
                )
                
                fig_empresas.update_traces(
                    textposition='inside', 
                    textinfo='percent',
                    insidetextorientation='horizontal', # Mantém o texto reto e legível
                    hovertemplate="<b>%{label}</b><br>Volume: R$ %{value:,.2f}<br>Representação: %{percent}"
                )
                
                fig_empresas.update_layout(
                    showlegend=False, 
                    margin=dict(l=20, r=20, t=50, b=20)
                )
                
                st.plotly_chart(fig_empresas, use_container_width=True)

                # A tabela completa para auditoria
                with st.expander("📋 Ver lista COMPLETA de valores por empresa"):
                    df_show_emp = df_empresas_geral.sort_values('v_liq', ascending=False).copy()
                    df_show_emp['v_liq'] = df_show_emp['v_liq'].apply(lambda x: f"R$ {x:,.2f}")
                    df_show_emp.rename(columns={'ose': 'Empresa', 'v_liq': 'Volume Total'}, inplace=True)
                    st.dataframe(df_show_emp, use_container_width=True, hide_index=True)

                # =======================================================
                # === BÔNUS: EVOLUÇÃO DE CUSTOS POR EMPRESA (FINAL) ===
                # =======================================================
                st.write("---")
                st.subheader("📈 Evolução de Custos por Empresa")

                col_comp = None
                for c in df_sec4.columns:
                    if "competencia" in c.lower().strip() or "comp" in c.lower().strip():
                        col_comp = c
                        break

                if not col_comp:
                    st.error("❌ Coluna de Competência não encontrada.")
                else:
                    lista_empresas = sorted(df_empresas_geral['ose'].unique().tolist())
                    empresa_selecionada = st.selectbox(
                        "Selecione a Empresa para análise detalhada:", 
                        lista_empresas, key="sel_evolucao_empresa"
                    )

                    if empresa_selecionada:
                        df_evol = df_sec4[df_sec4['ose'] == empresa_selecionada].copy()
                        
                        # --- TRATAMENTO DE DATAS ---
                        def tratar_data_flexivel(v):
                            v_str = str(v).strip().replace('.0', '')
                            if not v_str or v_str == 'nan': return pd.NaT
                            # Se for só o número, assume 2026 (ajuste se necessário para 2025)
                            if v_str.isdigit() and 1 <= int(v_str) <= 12:
                                return pd.to_datetime(f"{v_str.zfill(2)}/2026", format='%m/%Y')
                            return pd.to_datetime(v_str, errors='coerce', dayfirst=True)

                        df_evol['dt_ordem'] = df_evol[col_comp].apply(tratar_data_flexivel)
                        
                        # --- FILTRO 1: Remove o que não virou data e o que está no futuro (ex: Dez/26) ---
                        hoje = pd.Timestamp.now()
                        df_evol = df_evol[df_evol['dt_ordem'] <= hoje].dropna(subset=['dt_ordem'])
                        
                        if df_evol.empty:
                            st.warning(f"⚠️ Sem dados históricos processados para '{empresa_selecionada}'.")
                        else:
                            # Nome oficial solicitado pelo comando
                            df_evol['mes_competencia'] = df_evol['dt_ordem'].dt.strftime('%m/%Y')
                            
                            # Agrupamento real
                            df_evol_grafico = df_evol.groupby(['dt_ordem', 'mes_competencia'])['v_liq'].sum().reset_index()
                            df_evol_grafico = df_evol_grafico.sort_values('dt_ordem')

                            # --- GRÁFICO DE ÁREA MODERNO ---
                            fig_line = px.area(
                                df_evol_grafico, 
                                x='mes_competencia', 
                                y='v_liq',
                                title=f"Histórico de Custos: {empresa_selecionada}",
                                markers=True,
                                template="plotly_white"
                            )

                            fig_line.update_traces(
                                line_color='#2c5d71', 
                                fillcolor='rgba(44, 93, 113, 0.2)',
                                marker=dict(size=10, color='#2c5d71', symbol="circle", line=dict(width=2, color="white")),
                                hovertemplate="<b>%{x}</b><br>Volume: R$ %{y:,.2f}<extra></extra>"
                            )

                            fig_line.update_layout(
                                xaxis_title="Mês de Competência",
                                yaxis_title="Valor Líquido (R$)",
                                hovermode="x unified",
                                yaxis_tickprefix="R$ ",
                                margin=dict(l=20, r=20, t=60, b=20),
                                plot_bgcolor='rgba(0,0,0,0)',
                                paper_bgcolor='rgba(0,0,0,0)'
                            )

                            # --- FILTRO 2: Força o eixo X a ser categórico (Não inventa meses vazios) ---
                            fig_line.update_xaxes(type='category', showgrid=False)
                            fig_line.update_yaxes(showgrid=True, gridwidth=1, gridcolor='rgba(0,0,0,0.05)')

                            st.plotly_chart(fig_line, use_container_width=True)
                
           
            # =======================================================
            # === SEÇÃO 5: AGUARDANDO NOTA FISCAL (STATUS 6) ===
            # =======================================================
            st.divider()
            st.subheader("📝 5. Situação das Notas de Empenho aguardando Nota Fiscal")

            df_sec5 = df.copy()
            
            # Limpeza do status
            df_sec5['status_limpo'] = df_sec5['status'].astype(str).str.replace('.0', '', regex=False).str.strip()
            if 'v_liq' not in df_sec5.columns:
                df_sec5['v_liq'] = df_sec5['valor_liquido'].apply(limpar_valor)

            # Filtra apenas quem está no Status 6
            df_status6 = df_sec5[df_sec5['status_limpo'] == '6'].copy()

            if df_status6.empty:
                st.success("🎉 Excelente! Nenhuma fatura aguardando Nota Fiscal no momento.")
            else:
                df_status6['nup'] = df_status6['nup'].astype(str).str.strip()
                
                try:
                    aba_h_sec5 = sh.worksheet("SISAFA-NAVAL-historico")
                    df_hist_temp = pd.DataFrame(aba_h_sec5.get_all_records())
                    
                    if not df_hist_temp.empty:
                        df_hist_temp['nup'] = df_hist_temp['nup'].astype(str).str.strip()
                        df_hist_temp['status_destino'] = df_hist_temp['status_destino'].astype(str).str.replace('.0', '', regex=False).str.strip()
                        
                        df_hist_6 = df_hist_temp[df_hist_temp['status_destino'] == '6'].copy()
                        
                        if not df_hist_6.empty:
                            df_hist_6['data_hist'] = pd.to_datetime(df_hist_6['timestamp'], errors='coerce')
                            df_entrada_st6 = df_hist_6.groupby('nup')['data_hist'].max().reset_index()
                            df_status6 = df_status6.merge(df_entrada_st6, on='nup', how='left')
                except Exception as e:
                    st.toast("Aviso: Falha ao cruzar com o histórico.")

                # Fallback de segurança para as datas
                if 'data_hist' not in df_status6.columns:
                    df_status6['data_hist'] = pd.NaT
                
                col_data_proc = 'timestamp' if 'timestamp' in df_status6.columns else 'data_atualizacao' if 'data_atualizacao' in df_status6.columns else None
                
                if col_data_proc:
                    df_status6['data_proc'] = pd.to_datetime(df_status6[col_data_proc], errors='coerce')
                    df_status6['data_final'] = df_status6['data_hist'].fillna(df_status6['data_proc'])
                else:
                    df_status6['data_final'] = df_status6['data_hist']

                # Cálculo de dias
                hoje = pd.Timestamp.today()
                df_status6['dias_parada'] = (hoje - df_status6['data_final']).dt.days
                df_tempo = df_status6.dropna(subset=['dias_parada'])
                
                if not df_tempo.empty:
                    media_tempo = int(df_tempo['dias_parada'].mean())
                    max_tempo = int(df_tempo['dias_parada'].max())
                    min_tempo = int(df_tempo['dias_parada'].min())
                    idx_max = df_tempo['dias_parada'].idxmax()
                    empresa_antiga = df_tempo.loc[idx_max, 'ose']
                    nup_antigo = df_tempo.loc[idx_max, 'nup']
                else:
                    media_tempo, max_tempo, min_tempo = 0, 0, 0
                    empresa_antiga, nup_antigo = "N/A", "N/A"

                total_dinheiro_st6 = df_status6['v_liq'].sum()
                qtd_notas_st6 = len(df_status6) # Quantidade total de linhas no status 6

                # =======================================================
                # LAYOUT DE DESTAQUE (Financeiro + Quantidade)
                # =======================================================
                c_destaque1, c_destaque2 = st.columns(2)
                with c_destaque1:
                    st.markdown(f"### 💰 Volume Total: **R$ {total_dinheiro_st6:,.2f}**")
                with c_destaque2:
                    st.markdown(f"### 📑 Quantidade: **{qtd_notas_st6} Notas de Empenho**")
                
                st.write("") # Espaço

                # Linha de métricas de tempo
                c1, c2, c3 = st.columns(3)
                with c1:
                    st.metric(label="⏳ Média de Tempo", value=f"{media_tempo} dias", delta="Gargalo Médio", delta_color="inverse")
                with c2:
                    st.metric(label="⚠️ Fatura Mais Antiga", value=f"{max_tempo} dias", delta="💣💣💣💣💣", delta_color="inverse")
                    if max_tempo > 0:
                        st.caption(f"**NUP:** {nup_antigo}")
                        st.caption(f"🏢 {empresa_antiga}")
                with c3:
                    st.metric(label="😶‍🌫️ Fatura Mais Recente", value=f"{min_tempo} dias")

                # =======================================================
                # GRÁFICO GIGANTE (Sem legenda para maximizar tamanho)
                # =======================================================
                st.divider()
                
                df_emp_st6 = df_status6.groupby('ose')['v_liq'].sum().reset_index()
                
                fig_st6 = px.pie(
                    df_emp_st6, 
                    values='v_liq', 
                    names='ose', 
                    hole=0.45,
                    title="Distribuição Financeira (Passe o mouse para identificar a empresa)",
                    color_discrete_sequence=px.colors.qualitative.Safe
                )
                
                fig_st6.update_traces(
                    textposition='inside', 
                    textinfo='percent',
                    hovertemplate="<b>%{label}</b><br>Volume: R$ %{value:,.2f}<br>Representação: %{percent}"
                )
                fig_st6.update_layout(
                    showlegend=False, 
                    margin=dict(l=20, r=20, t=50, b=20)
                )
                st.plotly_chart(fig_st6, use_container_width=True)

                # =======================================================
                # =======================================================
                # NOVO: CARDS NEON DE COMPETÊNCIA (Distribuição Financeira)
                # =======================================================
                if 'mes_competencia' in df_status6.columns and 'ano_competencia' in df_status6.columns:
                    st.markdown("#### 🗓️ Montante aguardando emissão de NF por Competência")
                    
                    # Dicionário para deixar o mês com sigla padrão
                    meses_map_neon = {
                        1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR", 5: "MAI", 6: "JUN", 
                        7: "JUL", 8: "AGO", 9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ"
                    }
                    
                    # Função para juntar Mês e Ano bonitinho (ex: MAR/2026)
                    def montar_rotulo_comp(row):
                        try:
                            m = int(float(row['mes_competencia']))
                            a = int(float(row['ano_competencia']))
                            return f"{meses_map_neon.get(m, str(m))}/{a}"
                        except:
                            return f"{row['mes_competencia']}/{row['ano_competencia']}"
                    
                    # Cria a coluna temporária na base
                    df_status6['comp_neon'] = df_status6.apply(montar_rotulo_comp, axis=1)
                    
                    # Agrupa os valores financeiros pela nova coluna montada e ordena
                    df_comp_st6 = df_status6.groupby('comp_neon')['v_liq'].sum().reset_index()
                    df_comp_st6 = df_comp_st6.sort_values(by='v_liq', ascending=False)
                    
                    cores_neon = ['#00E676', '#2979FF', '#FFEA00', '#FF1744', '#D500F9']
                    
                    # TÁTICA BLINDADA: Lista para agrupar o HTML sem nenhuma quebra invisível
                    elementos_html = ['<div style="display: flex; flex-wrap: wrap; gap: 15px; margin-top: 10px; margin-bottom: 25px;">']
                    
                    for i, row in enumerate(df_comp_st6.itertuples()):
                        comp = str(row.comp_neon)
                        valor = float(row.v_liq)
                        cor = cores_neon[i % len(cores_neon)] 
                        
                        # O CARD INTEIRO EM UMA ÚNICA VARIÁVEL (Garante 0 erros de Markdown)
                        card = f'<div style="flex: 1; min-width: 150px; background-color: #1e1e1e; border: 1px solid {cor}; border-radius: 10px; padding: 15px; text-align: center; box-shadow: 0 0 15px {cor}40;"><div style="color: #cccccc; font-size: 13px; font-weight: bold; letter-spacing: 1px; margin-bottom: 5px;">{comp.upper()}</div><div style="color: {cor}; font-size: 19px; font-weight: 900; text-shadow: 0 0 10px {cor}80;">R$ {valor:,.2f}</div></div>'
                        
                        elementos_html.append(card)
                    
                    # FECHAMENTO OBRIGATÓRIO DA CAIXA (Era isso que estava faltando!)
                    elementos_html.append('</div>')
                    
                    # Une tudo num bloco só e injeta no Streamlit
                    html_final = "".join(elementos_html)
                    st.markdown(html_final, unsafe_allow_html=True)
                    
                else:
                    st.warning("⚠️ As colunas 'mes_competencia' e 'ano_competencia' não foram localizadas nesta base.")
                # =======================================================
                # =======================================================
                # =======================================================

                with st.expander("📋 Ver lista COMPLETA das NFs pendentes por empresa"):
                    df_show_st6 = df_emp_st6.sort_values('v_liq', ascending=False).copy()
                    df_show_st6['v_liq'] = df_show_st6['v_liq'].apply(lambda x: f"R$ {x:,.2f}")
                    df_show_st6.rename(columns={'ose': 'Empresa', 'v_liq': 'Volume Pendente'}, inplace=True)
                    st.dataframe(df_show_st6, use_container_width=True, hide_index=True)

            
            # =======================================================
            # === SEÇÃO 6: EVOLUÇÃO DOS VALORES AUDITADOS ===
            # =======================================================
            st.divider()
            st.subheader("📈 6. Evolução dos valores auditados")

            # 1. Filtro de Segurança e Preparação
            df_sec6 = df.copy()
            df_sec6['status_num'] = pd.to_numeric(df_sec6['status'], errors='coerce').fillna(0).astype(int)
            df_sec6 = df_sec6[df_sec6['status_num'] >= 3].copy() # Já auditados

            if df_sec6.empty:
                st.info("Aguardando faturas atingirem o status de Auditada (Status >= 3) para gerar o gráfico.")
            else:
                # Limpeza financeira
                df_sec6['v_liq'] = df_sec6['valor_liquido'].apply(limpar_valor)
                df_sec6['v_glosa'] = df_sec6['glosa'].apply(limpar_valor)

                # --- CATEGORIZAÇÃO RÍGIDA DO COMANDO ---
                cats_oficiais = ["1. OSE Civis", "2. HFA", "3. Base Op. Especiais (FUSEX)", "4. BAAN", "5. HFAB"]

                def categorizar_rigido(nome):
                    n = str(nome).upper()
                    if "HOSPITAL DAS FORÇAS ARMADAS" in n or "HFA" in n: return cats_oficiais[1]
                    elif "160098" in n or "OPERAÇÕES ESPECIAIS" in n: return cats_oficiais[2]
                    elif "120624" in n or "BASE AÉREA DE ANÁPOLIS" in n: return cats_oficiais[3]
                    elif "120096" in n or "HFAB" in n: return cats_oficiais[4]
                    return cats_oficiais[0]

                df_sec6['Categoria_Audit_Final'] = df_sec6['ose'].apply(categorizar_rigido)

                # --- PREPARAÇÃO CRONOLÓGICA ---
                df_sec6['sort_key'] = df_sec6['ano_competencia'] * 100 + df_sec6['mes_competencia']
                df_sec6['Competência'] = df_sec6.apply(
                    lambda x: f"{mapa_meses_abrev[int(x['mes_competencia'])]}/{str(int(x['ano_competencia']))[2:]}", axis=1
                )
                df_cronologico = df_sec6[['sort_key', 'Competência']].drop_duplicates().sort_values('sort_key')
                lista_competencias = df_cronologico['Competência'].tolist()

                # --- CONSTRUÇÃO DO GRÁFICO ---
                import plotly.graph_objects as go
                fig_audit = go.Figure()

                cores_audit = {
                    "1. OSE Civis": "#1abc9c", "2. HFA": "#1e3d33", 
                    "3. Base Op. Especiais": "#2e6b54", "4. BAAN Anápolis": "#529471", 
                    "5. HFAB": "#76b996", "Glosa Total": "#e74c3c"
                }

                ordem_empilhamento = ["1. OSE Civis", "5. HFAB", "4. BAAN Anápolis", "3. Base Op. Especiais", "2. HFA"]
                
                # Adicionar barras empilhadas (Valor Líquido)
                for cat in ordem_empilhamento:
                    y_vals = [df_sec6[(df_sec6['Categoria_Audit_Final'] == cat) & (df_sec6['Competência'] == comp)]['v_liq'].sum() for comp in lista_competencias]
                    fig_audit.add_trace(go.Bar(
                        name=cat, x=lista_competencias, y=y_vals, marker_color=cores_audit[cat], 
                        offsetgroup="Liquido", hovertemplate="<b>%{x}</b><br>%{data.name}: R$ %{y:,.2f}<extra></extra>"
                    ))

                # Adicionar barra Glosa
                y_glosa = [df_sec6[df_sec6['Competência'] == comp]['v_glosa'].sum() for comp in lista_competencias]
                fig_audit.add_trace(go.Bar(
                    name="Glosa Total", x=lista_competencias, y=y_glosa, marker_color=cores_audit["Glosa Total"], 
                    offsetgroup="Glosa", hovertemplate="<b>%{x}</b><br>Glosa Total: R$ %{y:,.2f}<extra></extra>"
                ))

                # --- LINHA DE TENDÊNCIA (Total Líquido) ---
                total_liq_mensal = [df_sec6[df_sec6['Competência'] == comp]['v_liq'].sum() for comp in lista_competencias]
                fig_audit.add_trace(go.Scatter(
                    x=lista_competencias, y=total_liq_mensal, name="Tendência", mode='lines+markers',
                    line=dict(color='black', width=3, dash='dot'), marker=dict(size=8, color='black')
                ))

                # --- REFINAMENTO DE ESTILO ---
                fig_audit.update_layout(
                    barmode='stack',
                    hovermode="x unified", paper_bgcolor='white', plot_bgcolor='white',
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
                    xaxis=dict(showgrid=False, linecolor='black', linewidth=1),
                    yaxis=dict(title="Montante (R$)", tickprefix="R$ ", gridcolor='rgba(0,0,0,0.05)')
                )

                st.plotly_chart(fig_audit, use_container_width=True)

            
            
            # =======================================================
            # =======================================================
            # === SEÇÃO 7: ANÁLISE DE CENTROS DE CUSTO (ORÇAMENTO) ===
            # =======================================================
            st.divider()
            st.subheader("🏢 7. Análise por Centro de Custo")

            try:
                # 1. Carga Segura dos Dados
                aba_auditoria = sh.worksheet("SISAFA-NAVAL-Auditoria")
                df_aud = pd.DataFrame(aba_auditoria.get_all_records())
                
                if df_aud.empty:
                    st.info("Aguardando inserção de dados na aba de Auditoria para carregar os Centros de Custo.")
                else:
                    # Padroniza cabeçalhos (remove espaços duplos)
                    df_aud.columns = df_aud.columns.astype(str).str.strip()
                    df_aud.columns = [" ".join(col.split()) for col in df_aud.columns]
                    
                    if 'nup' in df_aud.columns:
                        df_aud['nup'] = df_aud['nup'].astype(str).str.strip()
                    
                    # Criação da Chave Cronológica
                    df_aud['sort_key'] = df_aud['ano_competencia'] * 100 + df_aud['mes_competencia']
                    df_aud['Competência'] = df_aud.apply(
                        lambda x: f"{mapa_meses_abrev[int(x['mes_competencia'])]}/{str(int(x['ano_competencia']))[2:]}", axis=1
                    )
                    
                    # --- A MÁGICA DO "OUTROS" ---
                    for col in df_aud.columns:
                        if 'custo total' in col.lower():
                            df_aud.rename(columns={col: 'Outros'}, inplace=True)
                    
                    # --- LISTA OFICIAL RESTRITA DO COMANDO ---
                    lista_oficial = [
                        "Internações UTI (exceto OPME)", "Internações não UTI (exceto OPME)", "SIAD", "HOME CARE",
                        "Pequenas Cirurgias", "Consultas ambulatoriais", "Consultas emergenciais", "OPME",
                        "Remédio de Alto Custo: Quimioterápicos", "Remédio de Alto Custo: Imunobiológicos",
                        "Remédio de Alto Custo: Antibióticos", "Análises Clínicas", "RX Convencional", "Tomografias",
                        "Ressonâncias magnéticas", "Ultrassonografias", "Exames oftalmológicos", "Holter 24h", "Mapa 24h",
                        "Estudo eletrofisiológico (para estudo de arritmia cardíaca)", "Angiotomografia coronariana",
                        "Cintilografia miocárdica", "Teste Ergométrico", "Exames do Sistema Digestório e anexos",
                        "FACO (Catarata)", "Injeção Anti-VEGF (Ex: Lucentis)", "Revascularização miocárdica",
                        "Angioplastia coronariana com ou sem Stent", "Cateterismo cardíaco", "Hemodiálise", "Fisioterapia",
                        "Fonoaudiologia", "Psicologia / Psicoterapia", "Avaliação neuropsicológica", "Psicopedagogia",
                        "Terapia Ocupacional", "Musicoterapia", "Consultas", "Laboratórios Odontológicos",
                        "Ex. Radiol. e Doc. Orto", "Prótese", "Ortodontia", "Outros"
                    ]
                    
                    colunas_centros_custo = [col for col in lista_oficial if col in df_aud.columns]

                    # Função Blindada Financeira
                    def limpar_custo_auditoria(val):
                        if pd.isna(val) or val == '': return 0.0
                        if isinstance(val, (int, float)): return float(val)
                        v_str = str(val).replace('R$', '').strip()
                        if not v_str: return 0.0
                        
                        if '.' in v_str and ',' in v_str:
                            if v_str.rfind(',') > v_str.rfind('.'): 
                                v_str = v_str.replace('.', '').replace(',', '.')
                            else: 
                                v_str = v_str.replace(',', '')
                        elif ',' in v_str:
                            v_str = v_str.replace(',', '.') 
                        
                        try:
                            return float(v_str)
                        except:
                            return 0.0

                    for col in colunas_centros_custo:
                        df_aud[col] = df_aud[col].apply(limpar_custo_auditoria)

                    # Transformação Matricial (Melt)
                    df_long = pd.melt(
                        df_aud, 
                        id_vars=['Competência', 'sort_key', 'ose', 'nup'], 
                        value_vars=colunas_centros_custo,
                        var_name='Centro de Custo', 
                        value_name='Valor'
                    )
                    
                    df_long = df_long[df_long['Valor'] > 0].copy()

                    if df_long.empty:
                        st.warning("Nenhum lançamento financeiro maior que R$ 0,00 encontrado nos centros de custo oficiais.")
                    else:
                        # =======================================================
                        # VISÃO 1: GRÁFICO DE PIZZA MODERNO (COMPLETA)
                        # =======================================================
                        st.markdown("#### 📊 Distribuição Global de Gastos")
                        
                        df_pizza = df_long.groupby('Centro de Custo')['Valor'].sum().reset_index()
                        df_pizza = df_pizza.sort_values('Valor', ascending=False)
                        
                        fig_pie = px.pie(
                            df_pizza, 
                            names='Centro de Custo', 
                            values='Valor',
                            hole=0.45, 
                            color_discrete_sequence=px.colors.qualitative.Prism 
                        )
                        
                        fig_pie.update_traces(
                            textposition='inside', 
                            textinfo='percent',
                            hovertemplate="<b>%{label}</b><br>Gasto: R$ %{value:,.2f}<br>Representação: %{percent}<extra></extra>"
                        )
                        
                        fig_pie.update_layout(
                            paper_bgcolor='white', plot_bgcolor='white',
                            margin=dict(l=20, r=20, t=30, b=20),
                            legend=dict(font=dict(color='black'))
                        )
                        
                        st.plotly_chart(fig_pie, use_container_width=True)

                        
                        # =======================================================
                        # VISÃO 2: EVOLUÇÃO TEMPORAL - ÁREA COM SOMBRA (MODERNO)
                        # =======================================================
                        st.divider()
                        st.markdown("#### 📈 Evolução Individual por Centro de Custo")
                        
                        # Filtro Inteligente
                        lista_centros_ativos = sorted(df_long['Centro de Custo'].unique().tolist())
                        centro_selecionado = st.selectbox(
                            "🎯 Selecione a linha de serviço para visualização técnica:", 
                            lista_centros_ativos
                        )

                        # Filtra a base apenas para o centro selecionado
                        df_evol_individual = df_long[df_long['Centro de Custo'] == centro_selecionado]
                        df_evol_individual = df_evol_individual.groupby(['Competência', 'sort_key'])['Valor'].sum().reset_index()
                        df_evol_individual = df_evol_individual.sort_values('sort_key')

                        # Paleta tática para manter a cor consistente
                        cor_tatica = '#2c5d71' 

                        # Gráfico de Área Sombreado
                        fig_area_ind = px.area(
                            df_evol_individual, 
                            x='Competência', 
                            y='Valor', 
                            title=f"Histórico de Desembolso: {centro_selecionado}",
                            markers=True,
                            template="plotly_white"
                        )
                        
                        # Estilização "Bonitão" com sombra
                        fig_area_ind.update_traces(
                            line_color=cor_tatica, 
                            fillcolor='rgba(44, 93, 113, 0.2)', # Efeito sombreado suave
                            marker=dict(size=8, color=cor_tatica, line=dict(width=2, color="white")),
                            hovertemplate="<b>%{x}</b><br>Gasto: R$ %{y:,.2f}<extra></extra>"
                        )
                        
                        fig_area_ind.update_layout(
                            hovermode="x unified",
                            paper_bgcolor='white', plot_bgcolor='white',
                            margin=dict(l=20, r=20, t=50, b=20),
                            xaxis=dict(showgrid=False, type='category', linecolor='black', tickfont=dict(color='black')),
                            yaxis=dict(title="Total Gasto (R$)", tickprefix="R$ ", gridcolor='rgba(0,0,0,0.05)', tickfont=dict(color='black'))
                        )
                        
                        st.plotly_chart(fig_area_ind, use_container_width=True)

                        # =======================================================
                        # TABELA DE CONFERÊNCIA MATRICIAL
                        # =======================================================
                        with st.expander("📋 Ver Matriz Contábil Completa de Centros de Custo por Mês"):
                            df_matrix = df_long.pivot_table(
                                index='Centro de Custo', 
                                columns='Competência', 
                                values='Valor', 
                                aggfunc='sum'
                            ).fillna(0)
                            
                            df_matrix['Total Acumulado'] = df_matrix.sum(axis=1)
                            df_matrix = df_matrix.sort_values('Total Acumulado', ascending=False)
                            
                            st.dataframe(df_matrix.style.format("R$ {:,.2f}"), use_container_width=True)

            except Exception as e:
                st.error(f"Erro ao processar a inteligência de centros de custo: {e}")
        
        
        # =================================================================
        # 2. ABA: PRODUTIVIDADE (Médias Reais, Ciclo Total e Extremos)
        # =================================================================
        with tab_prod:
            st.subheader("🧑‍💻 Análise macroprocessual")
            st.subheader("⏱️ Análise Estatística de Permanência por Etapa (restrita às OSE civis)")
            
            try:
                aba_h = sh.worksheet("SISAFA-NAVAL-historico")
                df_hist_raw = pd.DataFrame(aba_h.get_all_records())
                
                if df_hist_raw.empty:
                    st.info("Aguardando dados históricos para análise.")
                else:
                    df_h = df_hist_raw.copy()
                    
                    # =================================================================
                    # 🔒 FILTRO DE DIRECIONAMENTO: ISOLANDO APENAS OSE CIVIS
                    # =================================================================
                    # 1. Carrega a Tabela A e filtra as OSEs do tipo CIVIL
                    aba_a = sh.worksheet("SISAFA-NAVAL-Tabela-A")
                    df_tabela_a = pd.DataFrame(aba_a.get_all_records())
                    nomes_civis = df_tabela_a[df_tabela_a['Tipo'] == 'CIVIL']['Razão Social'].tolist()

                    # 2. Descobre quais NUPs pertencem a essas OSEs na base principal (df)
                    nups_civis = df[df['ose'].isin(nomes_civis)]['nup'].astype(str).unique().tolist()
                    
                    # 1. Vacina de Tipos e Limpeza
                    df_h['nup'] = df_h['nup'].astype(str).str.strip()
                    col_dest = 'status_destinou' if 'status_destinou' in df_h.columns else 'status_destino'
                    df_h[col_dest] = df_h[col_dest].astype(str).str.replace('.0', '', regex=False).str.strip()
                    df_h['timestamp'] = pd.to_datetime(df_h['timestamp'], format='mixed', errors='coerce')
                    df_h = df_h.dropna(subset=['timestamp']).sort_values(['nup', 'timestamp'])
                    
                    # 3. APLICA O FILTRO CIVIL NO HISTÓRICO
                    df_h = df_h[df_h['nup'].isin(nups_civis)]

                    if df_h.empty:
                        st.warning("Sem dados históricos mapeados para OSEs do tipo CIVIL no momento.")
                    else:
                        # 2. Cálculo de Permanência Individual
                        df_h['data_saida'] = df_h.groupby('nup')['timestamp'].shift(-1)
                        df_h['data_saida'] = df_h['data_saida'].fillna(pd.Timestamp.now())
                        df_h['tempo_na_etapa'] = (df_h['data_saida'] - df_h['timestamp']).dt.total_seconds() / 86400
                        
                        # 3. Mapeamento das Etapas
                        mapa_status = {                
                            "1": "1. 📥 Cadastrada (SECOM)",
                            "2": "2. 🩺 Em Auditagem",
                            "3": "3. ✅ Auditada",
                            "4": "4. 💰 Aguardando NE",
                            "5": "5. 🏦 Empenhada",
                            "6": "6. 📝 Aguardando NF",
                            "7": "7. ⏳ Em liquidação",
                            "8": "8. 🖥️ Liquidada",
                        }
                        
                        df_h['nome_etapa'] = df_h[col_dest].map(mapa_status)

                        # 4. Criação da Tabela de Tempos
                        tabela_tempos = df_h.pivot_table(
                            index='nup', 
                            columns='nome_etapa', 
                            values='tempo_na_etapa', 
                            aggfunc='sum' 
                        )

                        # --- 5. CÁLCULO DAS ESTATÍSTICAS POR ETAPA ---
                        estatisticas_etapa = tabela_tempos.agg(['mean', 'min', 'max']).T.reset_index()
                        estatisticas_etapa.columns = ['Etapa', 'Média', 'Mínimo', 'Máximo']
                        
                        # Ordenação conforme o fluxo lógico
                        ordem_fluxo = list(mapa_status.values())
                        estatisticas_etapa['Etapa'] = pd.Categorical(estatisticas_etapa['Etapa'], categories=ordem_fluxo, ordered=True)
                        estatisticas_etapa = estatisticas_etapa.sort_values('Etapa')

                        # --- 6. GRÁFICO DE BARRAS (MÉDIAS REAIS) ---
                        fig_real = px.bar(
                            estatisticas_etapa, x='Etapa', y='Média',
                            color='Média', color_continuous_scale='Reds',
                            text_auto='.1f', title="Tempo Médio de Espera Real - OSE Civis (Dias)"
                        )
                        st.plotly_chart(fig_real, use_container_width=True)

                        # --- 7. MÉTRICAS CONSOLIDADAS ---
                        st.divider()
                        
                        tempo_ciclo_total = ...
                        tempo_ciclo_total = estatisticas_etapa['Média'].sum()
                        
                        c1, c2 = st.columns(2)
                        with c1:
                            st.metric("⏳ Tempo Médio Total do Ciclo", f"{tempo_ciclo_total:.1f} dias")
                            st.caption("Soma das médias reais das faturas Civis por todas as etapas.")
                        with c2:
                            # Taxa de liquidação calculada APENAS para o universo CIVIL
                            df_civis_main = df[df['nup'].isin(nups_civis)]
                            total_nup = df_civis_main['nup'].nunique()
                            concluidos = df_civis_main[df_civis_main['status'].astype(str).str.contains('8', na=False)]['nup'].nunique()
                            taxa = (concluidos / total_nup * 100) if total_nup > 0 else 0
                            st.metric("✅ Taxa de Liquidação (Civil)", f"{taxa:.1f}%")

                        # --- 8. TABELA DE EXTREMOS (Mín/Máx por Etapa) ---
                        st.write("### 🔍 Detalhamento de Performance por Etapa (Civil)")
                        df_formatado = estatisticas_etapa.copy()
                        for col in ['Média', 'Mínimo', 'Máximo']:
                            df_formatado[col] = df_formatado[col].apply(lambda x: f"{x:.2f} dias" if pd.notnull(x) else "-")
                        
                        st.table(df_formatado.set_index('Etapa'))

            except Exception as e:
                st.error(f"Erro na análise estatística: {e}")

            # =================================================================
            # 2. ABA: PRODUTIVIDADE (PM4PY + FILTRO DE MÊS)
            # =================================================================
            with tab_prod:
                st.header("🧭 Inteligência de Processos")
                
                try:
                    # 1. CARGA DE DADOS DO HISTÓRICO
                    aba_h = sh.worksheet("SISAFA-NAVAL-historico")
                    df_hist_p = pd.DataFrame(aba_h.get_all_records())
                    
                    if df_hist_p.empty:
                        st.info("Aguardando registros no histórico para iniciar a mineração.")
                    else:
                        # --- TRATAMENTO (VACINA TZ) ---
                        df_hist_p['timestamp'] = pd.to_datetime(df_hist_p['timestamp'], format='mixed', errors='coerce').dt.tz_localize(None)
                        df_hist_p = df_hist_p.dropna(subset=['timestamp'])
                        
                        # --- FILTRO POR MÊS ---
                        df_hist_p['mes_ano'] = df_hist_p['timestamp'].dt.strftime('%m/%Y')
                        lista_meses = sorted(df_hist_p['mes_ano'].unique(), 
                                             key=lambda x: pd.to_datetime(x, format='%m/%Y'), 
                                             reverse=True)
                        
                        st.write("### 🔍 Filtros de Análise")
                        mes_sel = st.selectbox("Selecione o mês:", ["Todos os Meses"] + lista_meses, key="prod_mes_filter")
                        
                        if mes_sel != "Todos os Meses":
                            df_hist_p = df_hist_p[df_hist_p['mes_ano'] == mes_sel]
                        
                        if df_hist_p.empty:
                            st.warning("Sem dados para este período.")
                        else:
                            # --- NOMES DAS ETAPAS ---
                            mapa_nomes = {
                                1: "1. CADASTRADA", 2: "2. EM AUDITAGEM", 3: "3. AUDITADA",
                                4: "4. AGUARDANDO NE", 5: "5. EMPENHADA", 6: "6. AGUARDANDO NF",
                                7: "7. EM LIQUIDAÇÃO", 8: "8. LIQUIDADA", 9: "9. PAGA"
                            }

                            df_pm = df_hist_p[['nup', 'status_destino', 'timestamp']].copy()
                            df_pm['status_destino'] = df_pm['status_destino'].map(mapa_nomes).fillna(df_pm['status_destino'])
                            
                            df_pm = df_pm.rename(columns={
                                'nup': 'case:concept:name', 
                                'status_destino': 'concept:name', 
                                'timestamp': 'time:timestamp'
                            }).sort_values(['case:concept:name', 'time:timestamp'])

                            # Process Mining
                            event_log = pm4py.format_dataframe(df_pm, case_id='case:concept:name', activity_key='concept:name', timestamp_key='time:timestamp')

                            # Tabela de Caminhos
                            from pm4py.statistics.traces.generic.log import case_statistics
                            var_stats = case_statistics.get_variant_statistics(event_log)
                            
                            var_list = []
                            total_n = df_pm['case:concept:name'].nunique()
                            for stat in var_stats:
                                var_list.append({
                                    "Fluxo Realizado": " ➔ ".join(stat['variant']),
                                    "Faturas": stat['count'],
                                    "%": round((stat['count'] / total_n) * 100, 1)
                                })
                            
                            st.write("#### 🔝 Principais Caminhos Detectados")
                            st.table(pd.DataFrame(var_list).sort_values("Faturas", ascending=False).head(5))

                            st.divider()

                            # Mapa Visual
                            st.write("#### 🗺️ Mineração de processos")
                            try:
                                import shutil
                                dot_exe = shutil.which("dot")
                                if dot_exe:
                                    os.environ["GRAPHVIZ_DOT"] = dot_exe
                                
                                dfg, sa, ea = pm4py.discover_dfg(event_log)
                                pm4py.save_vis_dfg(dfg, sa, ea, "mapa_prod.png")
                                st.image("mapa_prod.png", caption=f"Fluxo Minerado - {mes_sel}", use_container_width=True)
                            except Exception as e_vis:
                                st.warning("Visualização do mapa indisponível.")

                except Exception as e:
                    st.error(f"Erro no processamento PM4PY: {e}")



            # =================================================================
            # 🕵️ ANÁLISE MICROPROCESSUAL: CICLO DE PAGAMENTO (APENAS STATUS 9)
            # =================================================================
            st.divider()
            st.subheader("🕵️ Análise Microprocessual")
            st.markdown("### ⏱️ Monitoramento de Fluxo: OSE Civis (Processos Concluídos)")

            try:
                # 1. CARGA E MAPEAMENTO DE CIVIS
                aba_a = sh.worksheet("SISAFA-NAVAL-Tabela-A")
                df_tabela_a = pd.DataFrame(aba_a.get_all_records())
                nomes_civis = df_tabela_a[df_tabela_a['Tipo'] == 'CIVIL']['Razão Social'].tolist()

                df_mapeamento = df[df['ose'].isin(nomes_civis)][['nup', 'ose']].copy()
                nups_civis = df_mapeamento['nup'].astype(str).unique().tolist()

                # 2. FILTRO DE CONCLUSÃO (STATUS 9)
                df_hist_raw['status_destino'] = df_hist_raw['status_destino'].astype(str).str.replace('.0', '', regex=False).str.strip()
                nups_com_pagamento = df_hist_raw[df_hist_raw['status_destino'] == '9']['nup'].astype(str).unique().tolist()

                # Intersecção: Apenas NUPs CIVIS que já foram PAGOS (Status 9)
                nups_analise_final = list(set(nups_civis) & set(nups_com_pagamento))

                # 3. FILTRO DO HISTÓRICO
                df_civis = df_hist_raw.copy()
                df_civis['nup'] = df_civis['nup'].astype(str).str.strip()
                df_civis = df_civis[df_civis['nup'].isin(nups_analise_final)]
                
                df_civis['timestamp'] = pd.to_datetime(df_civis['timestamp'], format='mixed', errors='coerce')
                df_civis = df_civis.dropna(subset=['timestamp'])

                # Filtro de Mês/Evolução
                df_civis['mes_ano'] = df_civis['timestamp'].dt.strftime('%m/%Y')
                opcoes_meses = ["Todos"] + sorted(df_civis['mes_ano'].unique().tolist(), 
                                                 key=lambda x: pd.to_datetime(x, format='%m/%Y'), 
                                                 reverse=True)
                
                sel_mes_micro = st.selectbox("📅 Selecione o Período para Evolução:", opcoes_meses, key="filt_micro_evol")

                if sel_mes_micro != "Todos":
                    df_civis = df_civis[df_civis['mes_ano'] == sel_mes_micro]

                if df_civis.empty:
                    st.warning(f"Sem faturas concluídas para o período {sel_mes_micro}.")
                else:
                    # 4. CÁLCULO DE PERMANÊNCIA
                    df_civis = df_civis.sort_values(['nup', 'timestamp'])
                    df_civis['data_saida'] = df_civis.groupby('nup')['timestamp'].shift(-1)
                    df_civis['data_saida'] = df_civis['data_saida'].fillna(pd.Timestamp.now())
                    df_civis['dias'] = (df_civis['data_saida'] - df_civis['timestamp']).dt.total_seconds() / 86400

                    df_civis = df_civis.merge(df_mapeamento, on='nup', how='left')
                    df_pivot = df_civis.pivot_table(index=['nup', 'ose'], columns='status_destino', values='dias', aggfunc='sum').reset_index()
                    
                    for s in ['6', '7', '8']:
                        if s not in df_pivot.columns: df_pivot[s] = 0

                    # --- DEFINIÇÃO DE CORES SISAFA ---
                    cor_ideal = '#16A085'   # Verde SISAFA
                    cor_atencao = '#F39C12' # Amarelo/Laranja SISAFA
                    cor_critico = '#C0392B' # Vermelho SISAFA

                    # --- 1️⃣ LIQUIDAÇÃO (Fases 6 e 7) ---
                    df_pivot['tempo_liquidacao'] = df_pivot['6'].fillna(0) + df_pivot['7'].fillna(0)
                    m_liq = df_pivot['tempo_liquidacao'].mean()
                    min_liq = df_pivot.loc[df_pivot['tempo_liquidacao'].idxmin()]
                    max_liq = df_pivot.loc[df_pivot['tempo_liquidacao'].idxmax()]
                    qtd_total = len(df_pivot)

                    st.markdown("#### 1️⃣ Eficiência na Liquidação (do momento da entrega da Nota de Empenho aos fiscais de contrato à efetiva liquidação ⏳)")
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Média Global", f"{m_liq:.1f} dias")
                    c2.metric("Menor tempo 😎", f"{min_liq['tempo_liquidacao']:.1f} d", help=f"OSE: {min_liq['ose']}")
                    c3.metric("Maior tempo 🤯", f"{max_liq['tempo_liquidacao']:.1f} d", help=f"OSE: {max_liq['ose']}")
                    c4.metric("Qtd Processos", f"{qtd_total} faturas")

                    ideal_l = len(df_pivot[df_pivot['tempo_liquidacao'] <= 20])
                    atencao_l = len(df_pivot[(df_pivot['tempo_liquidacao'] > 20) & (df_pivot['tempo_liquidacao'] <= 30)])
                    critico_l = len(df_pivot[df_pivot['tempo_liquidacao'] > 30])

                    fig_liq = px.pie(
                        values=[ideal_l, atencao_l, critico_l],
                        names=['Ideal (≤20d)', 'Atenção (21-30d)', 'Crítico (>30d)'],
                        hole=0.6,
                        color=['Ideal (≤20d)', 'Atenção (21-30d)', 'Crítico (>30d)'],
                        color_discrete_map={'Ideal (≤20d)': cor_ideal, 'Atenção (21-30d)': cor_atencao, 'Crítico (>30d)': cor_critico}
                    )
                    fig_liq.update_traces(textposition='inside', textinfo='percent')
                    fig_liq.update_layout(margin=dict(t=20, b=20, l=0, r=0), height=350, showlegend=True,
                                          legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.05))
                    st.plotly_chart(fig_liq, use_container_width=True)

                    st.divider()

                    # --- 2️⃣ PAGAMENTO (Fase 8) ---
                    st.markdown("#### 2️⃣ Eficiência no Pagamento (do momento da liquidação no SIAFI ao efetivo pagamento⏳)")
                    m_pag = df_pivot['8'].mean()
                    min_pag = df_pivot.loc[df_pivot['8'].idxmin()]
                    max_pag = df_pivot.loc[df_pivot['8'].idxmax()]

                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Média Global", f"{m_pag:.1f} dias")
                    c2.metric("Menor tempo 😎", f"{min_pag['8']:.1f} d", help=f"OSE: {min_pag['ose']}")
                    c3.metric("Maior tempo 🤯", f"{max_pag['8']:.1f} d", help=f"OSE: {max_pag['ose']}")
                    c4.metric("Qtd Processos", f"{qtd_total} faturas")

                    ideal_p = len(df_pivot[df_pivot['8'] <= 3])
                    atencao_p = len(df_pivot[(df_pivot['8'] > 3) & (df_pivot['8'] <= 10)])
                    critico_p = len(df_pivot[df_pivot['8'] > 10])

                    fig_pag = px.pie(
                        values=[ideal_p, atencao_p, critico_p],
                        names=['Ideal (≤3d)', 'Atenção (4-10d)', 'Crítico (>10d)'],
                        hole=0.6,
                        color=['Ideal (≤3d)', 'Atenção (4-10d)', 'Crítico (>10d)'],
                        color_discrete_map={'Ideal (≤3d)': cor_ideal, 'Atenção (4-10d)': cor_atencao, 'Crítico (>10d)': cor_critico}
                    )
                    fig_pag.update_traces(textposition='inside', textinfo='percent')
                    fig_pag.update_layout(margin=dict(t=20, b=20, l=0, r=0), height=350, showlegend=True,
                                          legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.05))
                    st.plotly_chart(fig_pag, use_container_width=True)

            except Exception as e:
                st.error(f"Erro na análise microprocessual: {e}")

            #
            # =================================================================
            # --- 3️⃣ RANKING DE GESTORES (Fase 6 - Liquidação) ---
            # =================================================================
            st.divider()
            st.markdown("#### 3️⃣ Análise de Tempo Médio por Fiscal")
            st.info("💡 Exibe o tempo médio (em dias) que as faturas permanecem aguardando a Nota Fiscal sob a responsabilidade de cada Gestor Titular. 🐆")

            col_titular = 'Gestor Titular'

            if col_titular not in df_tabela_a.columns:
                st.warning("⚠️ Coluna 'Gestor Titular' não encontrada na Tabela A para gerar o ranking.")
            else:
                # 1. Criar um dataframe de mapeamento OSE -> Gestor Titular
                df_map_gestores = df_tabela_a[['Razão Social', col_titular]].rename(columns={'Razão Social': 'ose', col_titular: 'Gestor'})
                df_map_gestores = df_map_gestores.dropna(subset=['Gestor'])
                df_map_gestores['Gestor'] = df_map_gestores['Gestor'].astype(str).str.strip()
                df_map_gestores = df_map_gestores[df_map_gestores['Gestor'] != '']

                # 2. Fazer o cruzamento (merge) com a df_pivot (que já tem os tempos calculados)
                df_ranking_base = df_pivot.merge(df_map_gestores, on='ose', how='left')
                df_ranking_base['Gestor'] = df_ranking_base['Gestor'].fillna('Sem Gestor Mapeado')

                # =================================================================
                # 🔒 O FILTRO DE OURO: Apenas Ciclo 6 Completo (Saída garantida pro 7)
                # =================================================================
                # Identifica quais faturas efetivamente deram entrada no Status 7
                nups_que_chegaram_no_7 = df_civis[df_civis['status_destino'] == '7']['nup'].unique()

                # A fatura entra no cálculo SE: 
                # (i) Teve tempo rodando no status 6
                # (ii) O NUP está na lista dos que já alcançaram o status 7
                df_fase6 = df_ranking_base[
                    (df_ranking_base['6'] > 0) & 
                    (df_ranking_base['nup'].isin(nups_que_chegaram_no_7))
                ].copy()

                if df_fase6.empty:
                    st.info("Nenhum ciclo completo da Fase 6 (5 ➔ 6 ➔ 7) computado para o período.")
                else:
                    # 4. Agrupar, calcular as médias e métricas extremas
                    df_ranking = df_fase6.groupby('Gestor').agg(
                        Tempo_Medio=('6', 'mean'),
                        Menor_Tempo=('6', 'min'),
                        Maior_Tempo=('6', 'max'),
                        Qtd_Faturas=('6', 'count')
                    ).reset_index()

                    # Ordenar do Pior (Maior tempo) para o Melhor (Menor tempo)
                    df_ranking = df_ranking.sort_values(by='Tempo_Medio', ascending=False)

                    # Arredondar para 1 casa decimal
                    df_ranking['Tempo_Medio'] = df_ranking['Tempo_Medio'].round(1)
                    df_ranking['Menor_Tempo'] = df_ranking['Menor_Tempo'].round(1)
                    df_ranking['Maior_Tempo'] = df_ranking['Maior_Tempo'].round(1)

                    # --- 1. RENDERIZAÇÃO DO GRÁFICO (Tela Cheia) ---
                    fig_fiscais = px.bar(
                        df_ranking, 
                        x='Tempo_Medio', 
                        y='Gestor', 
                        orientation='h',
                        text='Tempo_Medio',
                        color='Tempo_Medio',
                        color_continuous_scale=[cor_ideal, cor_atencao, cor_critico], # Usa as cores do SISAFA
                        title="Tempo médio por Gestor (da entrega da NE ao (à) gestor (a) à entrega da NF para a Execução Financeira)"
                    )
                    
                    # CÁLCULO DE ALTURA DINÂMICA: 25 pixels de altura para cada Gestor na lista (mínimo de 400px)
                    altura_dinamica = max(400, len(df_ranking) * 25)

                    fig_fiscais.update_layout(
                        yaxis={'categoryorder':'total ascending'}, 
                        margin=dict(l=0, r=20, t=40, b=0),
                        coloraxis_showscale=False, # Esconde a legenda lateral
                        height=altura_dinamica # Aplica a altura calculada para não achatar
                    )
                    fig_fiscais.update_traces(textposition='outside')
                    
                    # Plota o gráfico ocupando 100% da largura
                    st.plotly_chart(fig_fiscais, use_container_width=True)

                    # --- 2. RENDERIZAÇÃO DA TABELA (Embaixo do Gráfico) ---
                    st.markdown("<br>**📊 Panorama Comparativo Detalhado**", unsafe_allow_html=True)
                    df_display = df_ranking.rename(columns={
                        'Tempo_Medio': 'Média (dias)',
                        'Maior_Tempo': 'Max (dias)',
                        'Menor_Tempo': 'Min (dias)',
                        'Qtd_Faturas': 'Faturas'
                    })
                    
                    # Tabela de suporte ocupando 100% da largura
                    st.dataframe(
                        df_display, 
                        use_container_width=True, 
                        hide_index=True,
                        column_config={
                            "Média (dias)": st.column_config.NumberColumn(format="%.1f ⏱️"),
                            "Max (dias)": st.column_config.NumberColumn(format="%.1f 🔴"),
                            "Min (dias)": st.column_config.NumberColumn(format="%.1f 🟢")
                        }
                    )

            
                    
                    # =================================================================
                    # =================================================================
                    # 🎯 RADAR 3D EXECUTIVO: CICLO STATUS 6 -> 7
                    # =================================================================
                    st.markdown("<br><br>🎯 Radar dos fiscais de contrato: gráfico de dispersão 📈", unsafe_allow_html=True)

                    # 1. Inicialização segura
                    df_radar = pd.DataFrame() 

                    try:
                        # Preparação básica dos dados
                        df_6 = df_civis[df_civis['status_destino'] == '6'].groupby('nup')['timestamp'].min().reset_index()
                        df_7 = df_civis[df_civis['status_destino'] == '7'].groupby('nup')['timestamp'].min().reset_index()
                        
                        df_radar = pd.merge(df_6, df_7, on='nup', suffixes=('_6', '_7'))
                        df_radar.rename(columns={'timestamp_6': 'Entrada_Fisc', 'timestamp_7': 'Retorno_Exec'}, inplace=True)
                        
                        # --- LOGS DE EMAIL ---
                        aba_logs = sh.worksheet("SISAFA-NAVAL-logs_acoes")
                        df_logs = pd.DataFrame(aba_logs.get_all_records())
                        df_logs.columns = df_logs.columns.str.strip()
                        df_logs['nup'] = df_logs['nup'].astype(str).str.strip()
                        
                        df_emails = df_logs[df_logs['acao'].str.strip().str.upper() == 'SOLICITACAO_NF_ENVIADA'].groupby('nup')['data_hora'].min().reset_index()
                        df_emails.rename(columns={'data_hora': 'Envio_Email'}, inplace=True)
                        
                        df_radar = pd.merge(df_radar, df_emails, on='nup', how='left')
                        df_radar['Envio_Email'] = df_radar['Envio_Email'].fillna("e-mail não enviado via SISAFA")

                        # Cruzamentos finais
                        df_mapa = df_tabela_a[['Razão Social', 'Gestor Titular', 'Gestor Substituto']].rename(columns={'Razão Social': 'ose'}).drop_duplicates()
                        df_radar = pd.merge(df_radar, df[['nup', 'ose', 'valor_liquido']], on='nup', how='left')
                        df_radar = pd.merge(df_radar, df_mapa, on='ose', how='left')
                        
                        # Cálculos
                        df_radar['Entrada_Fisc'] = pd.to_datetime(df_radar['Entrada_Fisc'])
                        df_radar['Retorno_Exec'] = pd.to_datetime(df_radar['Retorno_Exec'])
                        df_radar['Dias_Fisc'] = (df_radar['Retorno_Exec'] - df_radar['Entrada_Fisc']).dt.total_seconds() / 86400
                        df_radar['Valor_Num'] = df_radar['valor_liquido'].apply(limpar_valor)
                        
                        ordem_sla = ['Ideal👍', 'Atenção⚠️', 'Crítico💀']
                        df_radar['SLA'] = pd.cut(df_radar['Dias_Fisc'], bins=[-1, 20, 30, 9999], labels=ordem_sla)
                        df_radar['Tamanho_Bolha'] = df_radar['Valor_Num'].apply(lambda x: x if x > 1000 else 1000)
                        df_radar['Entrada_Fisc_Str'] = df_radar['Entrada_Fisc'].dt.strftime('%d/%m/%Y')
                        df_radar['Retorno_Exec_Str'] = df_radar['Retorno_Exec'].dt.strftime('%d/%m/%Y')

                    except Exception as e:
                        st.error(f"Erro ao processar dados do Radar: {e}")

                    # 2. SELETOR E RENDERIZAÇÃO
                    if not df_radar.empty:
                        lista_gestores = sorted(list(set(df_tabela_a['Gestor Titular'].dropna().unique().tolist() + 
                                                        df_tabela_a['Gestor Substituto'].dropna().unique().tolist())))
                        sel_gestor_radar = st.selectbox("Filtrar Radar por Fiscal:", ["TODOS (Visão Geral)"] + lista_gestores, key="radar_filtro")

                        if sel_gestor_radar != "TODOS (Visão Geral)":
                            empresas = df_tabela_a[(df_tabela_a['Gestor Titular'] == sel_gestor_radar) | 
                                                (df_tabela_a['Gestor Substituto'] == sel_gestor_radar)]['Razão Social'].unique()
                            df_radar_view = df_radar[df_radar['ose'].isin(empresas)].copy()
                        else:
                            df_radar_view = df_radar.copy()

                        if not df_radar_view.empty:
                            # Plotagem 3D Corrigida
                            fig_radar = px.scatter_3d(
                                df_radar_view, x='Valor_Num', y='SLA', z='Dias_Fisc',
                                color='SLA', size='Tamanho_Bolha', size_max=40,
                                category_orders={"SLA": ordem_sla},
                                color_discrete_map={ordem_sla[0]: '#2ecc71', ordem_sla[1]: '#f1c40f', ordem_sla[2]: '#e74c3c'},
                                hover_name='nup',
                                hover_data={
                                    'Valor_Num': False, 'SLA': False, 'Dias_Fisc': False,
                                    'Gestor Titular': True, 'Gestor Substituto': True, 'ose': True,
                                    'Entrada_Fisc_Str': True, 'Retorno_Exec_Str': True, 'Envio_Email': True
                                }
                            )

                            fig_radar.update_layout(
                                paper_bgcolor='white', plot_bgcolor='white', margin=dict(l=0, r=0, t=20, b=20),
                                showlegend=False,
                                scene=dict(
                                    aspectmode='manual', aspectratio=dict(x=1, y=0.5, z=1.5),
                                    xaxis=dict(title=dict(text="Valor líquido (R$)", font=dict(size=14)), backgroundcolor="white", gridcolor='lightgray', color='black'),
                                    yaxis=dict(title=dict(text="Status", font=dict(size=14)), backgroundcolor="white", gridcolor='lightgray', color='black', categoryorder='array', categoryarray=ordem_sla),
                                    zaxis=dict(title=dict(text="Período aguardando NF", font=dict(size=14)), backgroundcolor="white", gridcolor='lightgray', color='black'),
                                    bgcolor='white'
                                )
                            )
                            st.plotly_chart(fig_radar, use_container_width=True)
                        else:
                            st.info("Nenhum dado para o filtro selecionado.")
                    else:
                        st.warning("Não há processos que concluíram o ciclo 6->7 para exibir no Radar.")

            # =================================================================
            # =================================================================
            # 🕵️ VISÃO INDIVIDUAL DO GESTOR/FISCAL
            # =================================================================
            st.divider()
            st.subheader("👤 Painel Individual do Gestor")

            try:
                # 1. Nomes EXATOS das colunas conforme a sua SISAFA-NAVAL-Tabela-A
                col_titular = 'Gestor Titular'
                col_substituto = 'Gestor Substituto'

                if col_titular not in df_tabela_a.columns:
                    st.error(f"❌ A coluna '{col_titular}' não foi encontrada na Tabela A. Verifique se não há espaços acidentais no cabeçalho.")
                else:
                    # 2. Criar lista única e limpa de todos os militares envolvidos na gestão
                    lista_gestores = sorted(list(set(
                        df_tabela_a[col_titular].dropna().astype(str).unique().tolist() + 
                        (df_tabela_a[col_substituto].dropna().astype(str).unique().tolist() if col_substituto in df_tabela_a.columns else [])
                    )))
                            
                    # Remove sujeiras como campos vazios, 'nan' ou 'None'
                    lista_gestores = [g for g in lista_gestores if g.strip() and g.upper() not in ['NAN', 'NONE', '']]

                    sel_gestor = st.selectbox("Selecione o Fiscal:", [""] + lista_gestores, key="sel_gestor_individual")

                    if sel_gestor:
                        # --- FILTRAGEM DE CONTRATOS ---
                        contratos_titular = df_tabela_a[df_tabela_a[col_titular] == sel_gestor]
                        contratos_substituto = df_tabela_a[df_tabela_a[col_substituto] == sel_gestor] if col_substituto in df_tabela_a.columns else pd.DataFrame()

                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown(f"🏆 **Titularidade ({len(contratos_titular)})**")
                            if not contratos_titular.empty:
                                st.dataframe(contratos_titular[['Razão Social', 'Tipo', 'Termo de credenciamento']], use_container_width=True, hide_index=True)
                            else:
                                st.info("Nenhum contrato como titular.")

                        with c2:
                            st.markdown(f"🛡️ **Substituição ({len(contratos_substituto)})**")
                            if not contratos_substituto.empty:
                                st.dataframe(contratos_substituto[['Razão Social', 'Tipo', 'Termo de credenciamento']], use_container_width=True, hide_index=True)
                            else:
                                st.info("Nenhum contrato como substituto.")

                        # --- CÁLCULO FINANCEIRO (STATUS 6) ---
                        st.markdown(f"### 💰 Volume sob Responsabilidade do (a) gestor (a)")
                                
                        # Pegamos as empresas que ele é Titular
                        empresas_gestor = contratos_titular['Razão Social'].unique().tolist()
                                
                        # Filtramos os processos globais (df) que estão no Status 6 para essas empresas
                        df_status_6 = df[(df['status'] == 6) & (df['ose'].isin(empresas_gestor))].copy()
                                
                        if df_status_6.empty:
                            # Pega só o último nome para ficar amigável
                            nome_curto = sel_gestor.split()[-1] if sel_gestor else "Fiscal"
                            st.success(f"✅ Nada pendente no Status 6 para a carteira do(a) {nome_curto}! 🫡")
                        else:
                            df_status_6['valor_num'] = df_status_6['valor_apresentado'].apply(limpar_valor)
                                    
                            # Agrupamos por OSE para mostrar o montante por empresa
                            resumo_financeiro = df_status_6.groupby('ose').agg(
                                Qtd_Faturas=('nup', 'count'),
                                Total_Aberto=('valor_num', 'sum')
                            ).reset_index()

                            # Exibição
                            tot_gestor_s6 = resumo_financeiro['Total_Aberto'].sum()
                            st.metric(f"Total em Aberto (NF pendentes)", f"R$ {tot_gestor_s6:,.2f}")
                                    
                            st.write("**Detalhamento por OSE:**")
                            st.dataframe(
                                resumo_financeiro.rename(columns={'ose': 'Empresa', 'Qtd_Faturas': 'Faturas', 'Total_Aberto': 'Valor Total (R$)'}),
                                use_container_width=True,
                                hide_index=True,
                                column_config={
                                    "Valor Total (R$)": st.column_config.NumberColumn(format="R$ %.2f")
                                }
                            )

            except Exception as e:
                st.error(f"Erro ao carregar visão do gestor: {e}")




        # =================================================================
        # 3. ABA: ESTRUTURA DO SISAFA
        # =================================================================
        with tab_est:
            st.subheader("📂 Estrutura e Documentação")
            st.info("Esta seção contém a documentação técnica e os manuais do SISAFA-NAVAL.")
            st.write("---")
            st.markdown("**(Módulo em desenvolvimento - Aguardando upload dos arquivos técnicos)**")

        


    elif st.session_state.modulo_ativo == "OSE":
        st.header("🏥 Portal da OSE")
        
        # --- 1. PREPARAÇÃO DOS DADOS (A VACINA DOS 14 DÍGITOS) ---
        user_cnpj = str(st.session_state.user_id).strip().zfill(14)

        # Mapas de Tradução (Garantindo que existam no escopo)
        mapa_status_fisc = {
            1: "1 - FATURA CADASTRADA", 2: "2 - EM AUDITAGEM", 3: "3 - AUDITADA",
            4: "4 - AGUARDANDO EMISSÃO DE NE", 5: "5 - FATURA EMPENHADA",
            6: "6 - AGUARDANDO EMISSÃO DE NF", 7: "7 - EM LIQUIDAÇÃO",
            8: "8 - FATURA LIQUIDADA", 9: "9 - FATURA PAGA"
        }

        cores_map = {
            "1 - FATURA CADASTRADA": "#95a5a6", "2 - EM AUDITAGEM": "#f39c12",
            "3 - AUDITADA": "#3498db", "4 - AGUARDANDO EMISSÃO DE NE": "#f1c40f",
            "5 - FATURA EMPENHADA": "#9b59b6", "6 - AGUARDANDO EMISSÃO DE NF": "#e67e22",
            "7 - EM LIQUIDAÇÃO": "#e74c3c", "8 - FATURA LIQUIDADA": "#1abc9c",
            "9 - FATURA PAGA": "#27ae60"
        }

        # Preparando a Tabela-A (Dados do Fiscal)
        df_tabela_a = carregar_dados_cache(ABA_TABELA_A)
        df_tabela_a.columns = [c.strip().replace(' ', '_').upper() for c in df_tabela_a.columns]
        
        df_tabela_a = df_tabela_a.loc[:, ~df_tabela_a.columns.duplicated()]

        if 'CNPJ' in df_tabela_a.columns:
            # 3. Criamos o CNPJ_LIMPO com garantia de ser uma Series (coluna única)
            df_tabela_a['CNPJ_LIMPO'] = (
                df_tabela_a['CNPJ']
                .astype(str)
                .str.split('.').str[0]
                .str.strip()
                .str.zfill(14)
            )
            dados_minha_ose = df_tabela_a[df_tabela_a['CNPJ_LIMPO'] == user_cnpj].copy()
        else:
            st.error("❌ Coluna 'CNPJ' não localizada na Tabela A.")
            dados_minha_ose = pd.DataFrame()

        # --- Preparando as Faturas (df principal) ---
        if 'cnpj' in df.columns:
            # Aplicamos a mesma lógica de limpeza aqui por segurança
            df['cnpj_limpo'] = df['cnpj'].astype(str).str.split('.').str[0].str.strip().str.zfill(14)
            df_minhas_faturas = df[df['cnpj_limpo'] == user_cnpj].copy()
        else:
            df_minhas_faturas = pd.DataFrame()

        # --- 2. INTERFACE DAS ABAS ---
        tab_visao, tab_rel = st.tabs(["🔭 Visão Geral", "💬 Relacionamento"])

        # --- 1. ABA: VISÃO GERAL ---
        with tab_visao:
            # --- IMAGEM FIXA (RODAPÉ) ---
            mapeamento_path = carregar_imagem(caminho_mapeamento)
            if mapeamento_path:
                with open(mapeamento_path, "rb") as f:
                    data = base64.b64encode(f.read()).decode()
                    st.markdown(
                        f'<img src="data:image/png;base64,{data}" '
                        'style="position: fixed; bottom: 20px; right: 20px; width: 220px; z-index:998; opacity: 0.9; pointer-events: none;">',
                        unsafe_allow_html=True
                    )

            # =========================================================
            # 1. SEÇÃO FISCAL (AGORA COM E-MAILS E NIP)
            # =========================================================
            st.subheader("👮 Fiscais do Contrato")
            if not dados_minha_ose.empty:
                info = dados_minha_ose.iloc[0]
                c1, c2, c3 = st.columns(3)
                with c1:
                    st.markdown("**Fiscal Titular**")
                    st.write(f"👤 {info.get('GESTOR_TITULAR', 'Não informado')}")
                    # Tenta buscar o e-mail em diferentes variações de nome de coluna
                    email_t = info.get('E-MAIL_DO_GESTOR_TITULAR', info.get('EMAIL_TITULAR', 'E-mail não cadastrado'))
                    st.caption(f"📧 {email_t}")
                with c2:
                    st.markdown("**Fiscal Substituto**")
                    st.write(f"👤 {info.get('GESTOR_SUBSTITUTO', 'Não informado')}")
                    email_s = info.get('E-MAIL_DO_GESTOR_SUBSTITUTO', info.get('EMAIL_SUBSTITUTO', 'E-mail não cadastrado'))
                    st.caption(f"📧 {email_s}")
                
            else:
                st.info("Informações dos fiscais ainda não vinculadas.")

            st.divider()

            # =========================================================
            # 2. PREPARAÇÃO E LIMPEZA DE DADOS (VACINA V6 NUCLEAR)
            # =========================================================
            if df_minhas_faturas.empty:
                st.warning(f"Nenhuma fatura encontrada para o CNPJ: {user_cnpj}")
            else:
                def limpar_v6_nuclear(valor):
                    if pd.isna(valor) or str(valor).strip() in ["", "-", "None", "nan"]:
                        return 0.0
                    
                    # Se já for número real, não mexe
                    if isinstance(valor, (int, float)):
                        return float(valor)
                    
                    # 1. Limpeza bruta: remove R$ e espaços
                    s = str(valor).replace("R$", "").replace(" ", "").strip()
                    
                    # 2. Identifica o separador decimal (o último sinal que aparece)
                    ponto = s.rfind('.')
                    virgula = s.rfind(',')
                    
                    if virgula != -1 and ponto != -1:
                        if virgula > ponto: 
                            # Padrão BR (1.234,56): remove ponto, vira 1234,56, troca vírgula por ponto
                            s = s.replace(".", "").replace(",", ".")
                        else: 
                            # Padrão US/HFA (106,357.02): remove TODAS as vírgulas
                            s = s.replace(",", "")
                    elif virgula != -1:
                        # Se só tem vírgula, checa se é milhar (1,000) ou decimal (1,73)
                        # Em finanças, se só tem uma vírgula e 2 casas depois, é decimal.
                        if s.count(',') == 1 and len(s.split(',')[1]) == 2:
                            s = s.replace(",", ".")
                        else:
                            s = s.replace(",", "")
                    
                    try:
                        return float(s)
                    except:
                        return 0.0

                # Aplicamos a limpeza pesada
                cols_fin = ['valor_apresentado', 'valor_glosa', 'valor_liquido']
                for col in cols_fin:
                    if col in df_minhas_faturas.columns:
                        df_minhas_faturas[col] = df_minhas_faturas[col].apply(limpar_v6_nuclear)

                # =========================================================
                # 3. DASHBOARD FINANCEIRO (MÉTRICAS E PREPARAÇÃO)
                # =========================================================
                
                # --- 1. DEFINIÇÃO DOS MAPAS (Nomes e Meses) ---
                mapa_nomes = {
                    1: "1. CADASTRADA", 2: "2. EM AUDITAGEM", 3: "3. AUDITADA",
                    4: "4. AGUARDANDO NE", 5: "5. EMPENHADA", 6: "6. AGUARDANDO NF",
                    7: "7. EM LIQUIDAÇÃO", 8: "8. LIQUIDADA", 9: "9. PAGA"
                }

                mapa_meses_local = {
                    1: 'JAN', 2: 'FEV', 3: 'MAR', 4: 'ABR', 5: 'MAI', 6: 'JUN',
                    7: 'JUL', 8: 'AGO', 9: 'SET', 10: 'OUT', 11: 'NOV', 12: 'DEZ'
                }

                # --- 2. TRADUÇÃO DE STATUS (Cria a coluna Situação antes de tudo) ---
                if 'status' in df_minhas_faturas.columns:
                    df_minhas_faturas['Situação'] = df_minhas_faturas['status'].map(mapa_nomes).fillna("N/A")

                # Preparação do mês (sigla)
                if 'mes_sigla' not in df_minhas_faturas.columns and 'mes_competencia' in df_minhas_faturas.columns:
                    df_minhas_faturas['mes_sigla'] = pd.to_numeric(df_minhas_faturas['mes_competencia'], errors='coerce').map(mapa_meses_local)
                
                if 'mes_sigla' in df_minhas_faturas.columns:
                    df_minhas_faturas['mes_sigla'] = df_minhas_faturas['mes_sigla'].fillna("N/A")

                # Cálculos de Valores
                v_proc = df_minhas_faturas[df_minhas_faturas['status'] < 9]['valor_liquido'].sum()
                v_pago = df_minhas_faturas[df_minhas_faturas['status'] == 9]['valor_liquido'].sum()

                st.markdown(f"### 💰 Resumo Financeiro")
                m1, m2 = st.columns(2)
                
                with m1:
                    val_proc_fmt = f"R$ {v_proc:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    st.metric("Total em Aberto", val_proc_fmt)
                
                with m2:
                    val_pago_fmt = f"R$ {v_pago:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    st.metric("Total Pago (Liquidado)", val_pago_fmt)

                st.divider()

                # =========================================================
                # 4. TABELA DETALHADA
                # =========================================================
                st.subheader("📑 Detalhamento das Faturas")
                
                mapa_cols_exibicao = {
                    'Numero_da_fatura': 'Nº da fatura', 
                    'valor_apresentado': 'Valor Apresentado',
                    'valor_glosa': 'Glosa', 
                    'valor_liquido': 'Valor líquido',
                    'mes_sigla': 'Mês Competência', 
                    'ano_competencia': 'Ano',
                    'ne': 'NE', 'nf': 'NF', 'ob': 'OB', 
                    'Situação': 'Situação da Fatura' # <-- Agora ele encontra essa coluna!
                }
                
                cols_reais = [c for c in mapa_cols_exibicao.keys() if c in df_minhas_faturas.columns]
                
                if not df_minhas_faturas.empty:
                    df_final = df_minhas_faturas[cols_reais].copy()
                    if 'ano_competencia' in df_final.columns:
                        df_final = df_final.sort_values(by=['ano_competencia'], ascending=False)
                    
                    df_final = df_final.rename(columns=mapa_cols_exibicao)

                    formatos = {col: 'R$ {:,.2f}' for col in ['Valor Apresentado', 'Valor líquido', 'Glosa'] if col in df_final.columns}
                    st.dataframe(df_final.style.format(formatos), use_container_width=True, hide_index=True)
                else:
                    st.info("Nenhuma fatura encontrada.")

                # =========================================================
                # 5. GRÁFICO DE PIZZA (AO FINAL DA PÁGINA)
                # =========================================================
                st.write("") 
                if not df_minhas_faturas.empty:
                    import plotly.express as px
                    
                    st.markdown("#### 📊 Situação das faturas")
                    
                    # Agrupamento dinâmico pela Situação (Texto)
                    df_status_grafico = df_minhas_faturas.groupby('Situação')['valor_liquido'].sum().reset_index()
                    
                    fig = px.pie(
                        df_status_grafico, 
                        values='valor_liquido', 
                        names='Situação',
                        hole=0.4,
                        color_discrete_sequence=px.colors.qualitative.Prism 
                    )
                    
                    fig.update_layout(
                        height=400,
                        margin=dict(t=30, b=0, l=0, r=0),
                        legend=dict(orientation="h", yanchor="bottom", y=-0.3, xanchor="center", x=0.5)
                    )
                    
                    st.plotly_chart(fig, use_container_width=True)


        # --- 2. ABA: RELACIONAMENTO (PORTAL OSE) ---
        with tab_rel:
            st.subheader("💬 Central de Relacionamento")
            st.markdown("Utilize este espaço para tirar dúvidas sobre faturas específicas.")

            if df_minhas_faturas.empty:
                st.info("Você precisa ter faturas cadastradas para iniciar um contato.")
            else:
                # Criamos a etiqueta amigável para a OSE escolher a fatura
                df_minhas_faturas['label_fatura'] = (
                    "Fatura: " + df_minhas_faturas['Numero_da_fatura'].astype(str) + 
                    "/" + df_minhas_faturas['ano_competencia'].astype(str) + 
                    " (" + df_minhas_faturas['mes_sigla'] + ")"
                )
                
                fatura_sel_label = st.selectbox("Sobre qual fatura deseja falar?", [""] + df_minhas_faturas['label_fatura'].tolist())

                if fatura_sel_label:
                    dados_f = df_minhas_faturas[df_minhas_faturas['label_fatura'] == fatura_sel_label].iloc[0]
                    
                    # --- LÓGICA DE ROTEAMENTO AUTOMÁTICO ---
                    # Define para qual setor a dúvida vai de acordo com o status atual
                    status_atual = int(dados_f['status'])
                    if status_atual in [2, 3]:
                        setor_destino = "AUDITORIA"
                    elif status_atual in [4, 5, 7, 8, 9]:
                        setor_destino = "FINANCEIRO"
                    else:
                        setor_destino = "FISCALIZAÇÃO" # Caso esteja em outro status

                    st.info(f"📌 **Contexto:** {fatura_sel_label} | **Destino:** {setor_destino}")

                    with st.container(border=True):
                        assunto = st.text_input("Assunto da dúvida:", placeholder="Ex: Recurso de Glosa")
                        mensagem_texto = st.text_area("Descreva sua solicitação:")
                        
                        if st.button("📤 ENVIAR MENSAGEM OFICIAL", use_container_width=True):
                            if assunto and mensagem_texto:
                                try:
                                    aba_msg = sh.worksheet(ABA_MENSAGENS)
                                    
                                    # --- MONTAGEM DA LINHA (10 COLUNAS EXATAS) ---
                                    # Ordem: id_mensagem, nup, Numero_da_fatura, remetente, setor_destino, texto, data_envio, data_resposta, status_msg, respondido_por_nip
                                    nova_msg = [
                                        str(int(time.time())),           # 1. id_mensagem (Usamos o timestamp como ID único)
                                        dados_f['nup'],                  # 2. nup
                                        dados_f['Numero_da_fatura'],     # 3. Numero_da_fatura
                                        user_cnpj,                       # 4. remetente (CNPJ da OSE)
                                        setor_destino,                   # 5. setor_destino
                                        f"[{assunto}] {mensagem_texto}", # 6. texto (Juntei o assunto no corpo)
                                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"), # 7. data_envio
                                        "",                              # 8. data_resposta (Vazio)
                                        "PENDENTE",                      # 9. status_msg
                                        ""                               # 10. respondido_por_nip (Vazio)
                                    ]
                                    
                                    aba_msg.append_row(nova_msg)
                                    
                                    # Registro no Log Geral (Tabela-B)
                                    registrar_acao(dados_f['nup'], dados_f['Numero_da_fatura'], "CONTATO_OSE", f"Setor: {setor_destino}")
                                    
                                    st.success(f"Sua mensagem foi enviada para a {setor_destino} do HNBra!")
                                    time.sleep(1.5)
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Erro ao salvar na planilha: {e}")
                            else:
                                st.warning("Preencha o assunto e a mensagem.")

    # =========================================================================
    # ========================== MÓDULO ADMIN =================================
    # =========================================================================
    elif st.session_state.modulo_ativo == "ADMIN":
        st.header("⚙️ Painel de Administração do SISAFA")
        st.markdown("Gestão centralizada de usuários e cadastro de contratos. **A exclusão de registros é bloqueada por segurança.**")
        st.divider()

        tab_usuarios, tab_contratos, tab_estatisticas = st.tabs(["👥 Gestão de Usuários", "📄 Tabela de Contratos", "📊 Visão Global"])

        # -------------------------------------------------------------------------
        # ABA 1: GESTÃO DE USUÁRIOS
        # -------------------------------------------------------------------------
        with tab_usuarios:
            st.subheader("Controle de Acesso e Perfis")
            
            try:
                aba_usuarios = sh.worksheet("SISAFA-NAVAL-Usuarios")
                df_usuarios = pd.DataFrame(aba_usuarios.get_all_records())
                
                # Garantindo que os tipos de dados estão corretos para edição
                df_usuarios['NIP'] = df_usuarios['NIP'].astype(str)
                df_usuarios['Senha'] = df_usuarios['Senha'].astype(str)
                
                perfis_permitidos = ["Execução Financeira", "ADMIN", "Gerencial", "Fiscalização de Contrato", "SECOM", "Auditoria", "OSE", "FISCAL_GLOBAL"]
                
                with st.form("form_add_user"):
                    st.markdown("#### ➕ Incluir Novo Usuário / OSE")
                    c1, c2 = st.columns(2)
                    with c1:
                        novo_nip_cnpj = st.text_input("NIP ou CNPJ (Somente números)", help="Para militares use o NIP. Para OSE use o CNPJ (14 dígitos).")
                        novo_nome = st.text_input("Nome / Razão Social")
                        novo_perfil = st.selectbox("Perfil de Acesso", perfis_permitidos)
                    with c2:
                        novo_email = st.text_input("E-mail corporativo")
                        nova_senha = st.text_input("Senha Inicial", type="password")
                        confirmar_senha = st.text_input("Confirmar Senha", type="password")
                        
                    btn_add_user = st.form_submit_button("Cadastrar Usuário", use_container_width=True)
                    
                    if btn_add_user:
                        if not novo_nip_cnpj or not novo_nome or not nova_senha:
                            st.error("Preencha NIP/CNPJ, Nome e Senha obrigatoriamente.")
                        elif nova_senha != confirmar_senha:
                            st.error("As senhas não coincidem.")
                        elif novo_nip_cnpj in df_usuarios['NIP'].values:
                            st.error("Este NIP/CNPJ já está cadastrado no sistema!")
                        else:
                            # Adiciona no final da planilha
                            aba_usuarios.append_row([novo_nip_cnpj, novo_nome, novo_perfil, novo_email, nova_senha])
                            st.success(f"Usuário {novo_nome} cadastrado com sucesso!")
                            time.sleep(1.5)
                            st.rerun()
                
                st.write("")
                st.markdown("#### ✏️ Edição em Massa (Reset de Senha e Troca de Perfil)")
                st.info("Edite as células abaixo (dê dois cliques). Para salvar as alterações, clique no botão ao final da tabela. Adições de novas linhas pela tabela também são permitidas.")
                
                # O editor permite editar e adicionar, mas bloqueamos a exclusão (num_rows="dynamic" com validação no save)
                df_usuarios_editado = st.data_editor(
                    df_usuarios,
                    num_rows="dynamic",
                    use_container_width=True,
                    column_config={
                        "NIP": st.column_config.TextColumn("NIP / CNPJ", required=True),
                        "PERFIL": st.column_config.SelectboxColumn("Perfil", options=perfis_permitidos, required=True),
                        "Senha": st.column_config.TextColumn("Senha (Editável)", required=True)
                    },
                    hide_index=True,
                    key="editor_usuarios"
                )
                
                if st.button("💾 Salvar Alterações de Usuários", type="primary"):
                    if len(df_usuarios_editado) < len(df_usuarios):
                        st.error("⚠️ Operação Negada! A exclusão de usuários é bloqueada por segurança. Restaure a linha excluída ou recarregue a página.")
                    else:
                        aba_usuarios.clear()
                        # Monta o cabeçalho
                        aba_usuarios.update([df_usuarios_editado.columns.values.tolist()] + df_usuarios_editado.values.tolist())
                        st.success("Tabela de Usuários atualizada com sucesso no banco de dados!")
                        time.sleep(1.5)
                        st.rerun()

            except Exception as e:
                st.error(f"Erro ao acessar a tabela de usuários: {e}")

        # -------------------------------------------------------------------------
        # ABA 2: TABELA DE CONTRATOS (TABELA A)
        # -------------------------------------------------------------------------
        with tab_contratos:
            st.subheader("Gestão de Credenciamentos e Fiscais")
            try:
                aba_tabela_a = sh.worksheet("SISAFA-NAVAL-Tabela-A")
                df_tab_a = pd.DataFrame(aba_tabela_a.get_all_records())
                
                # Padronizando CNPJ e NIPs como texto para não perder zeros à esquerda
                for col in ['CNPJ', 'NIP do Gestor Titular', 'NIP do Gestor Substituto']:
                    if col in df_tab_a.columns:
                        df_tab_a[col] = df_tab_a[col].astype(str).str.replace('.0', '', regex=False)

                st.info("Altere as informações dos Fiscais, E-mails e dados de Edital diretamente na tabela abaixo. A exclusão de linhas não será aceita no salvamento.")
                
                df_tab_a_editado = st.data_editor(
                    df_tab_a,
                    num_rows="dynamic",
                    use_container_width=True,
                    hide_index=True,
                    key="editor_tabela_a"
                )
                
                if st.button("💾 Salvar Alterações da Tabela de Contratos", type="primary"):
                    if len(df_tab_a_editado) < len(df_tab_a):
                        st.error("⚠️ Operação Negada! A exclusão de contratos é bloqueada para manter o histórico financeiro intacto.")
                    else:
                        aba_tabela_a.clear()
                        aba_tabela_a.update([df_tab_a_editado.columns.values.tolist()] + df_tab_a_editado.values.tolist())
                        st.success("Tabela de Contratos atualizada com sucesso no banco de dados!")
                        time.sleep(1.5)
                        st.rerun()
                        
            except Exception as e:
                st.error(f"Erro ao acessar a Tabela A: {e}")

        # -------------------------------------------------------------------------
        # ABA 3: ESTATÍSTICAS E GRÁFICOS
        # -------------------------------------------------------------------------
        with tab_estatisticas:
            st.subheader("Panorama de Fiscalização")
            
            try:
                if 'df_tab_a' in locals() and not df_tab_a.empty:
                    import plotly.express as px
                    
                    c1, c2 = st.columns(2)
                    
                    # Gráfico 1: Gestores Titulares
                    with c1:
                        if 'Gestor Titular' in df_tab_a.columns:
                            df_titulares = df_tab_a['Gestor Titular'].value_counts().reset_index()
                            df_titulares.columns = ['Fiscal Titular', 'Quantidade de Contratos']
                            
                            # Remove os vazios/não informados para não sujar o gráfico
                            df_titulares = df_titulares[df_titulares['Fiscal Titular'].str.strip() != ""]
                            
                            fig_tit = px.pie(
                                df_titulares, 
                                values='Quantidade de Contratos', 
                                names='Fiscal Titular',
                                title="Distribuição de Contratos (Titulares)",
                                hole=0.4,
                                color_discrete_sequence=px.colors.sequential.Teal
                            )
                            fig_tit.update_traces(textposition='inside', textinfo='percent+value')
                            st.plotly_chart(fig_tit, use_container_width=True)
                        else:
                            st.warning("Coluna 'Gestor Titular' não encontrada.")

                    # Gráfico 2: Gestores Substitutos
                    with c2:
                        if 'Gestor Substituto' in df_tab_a.columns:
                            df_subs = df_tab_a['Gestor Substituto'].value_counts().reset_index()
                            df_subs.columns = ['Fiscal Substituto', 'Quantidade de Contratos']
                            
                            df_subs = df_subs[df_subs['Fiscal Substituto'].str.strip() != ""]
                            
                            fig_sub = px.pie(
                                df_subs, 
                                values='Quantidade de Contratos', 
                                names='Fiscal Substituto',
                                title="Distribuição de Contratos (Substitutos)",
                                hole=0.4,
                                color_discrete_sequence=px.colors.sequential.Burg
                            )
                            fig_sub.update_traces(textposition='inside', textinfo='percent+value')
                            st.plotly_chart(fig_sub, use_container_width=True)
                        else:
                            st.warning("Coluna 'Gestor Substituto' não encontrada.")
                else:
                    st.info("Acesse a aba 'Tabela de Contratos' primeiro para carregar os dados gráficos.")
                    
            except Exception as e:
                st.error(f"Erro ao gerar gráficos: {e}")
