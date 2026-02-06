import streamlit as st
from st_gsheets_connection import GSheetsConnection
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF
from num2words import num2words
from datetime import datetime, date, timedelta, time
import io
import urllib.parse

# ==============================================================================
# 1. CONFIGURAÇÃO E CONEXÃO
# ==============================================================================
st.set_page_config(page_title="Gerador de Ata SSVP (Cloud)", layout="wide", page_icon="✝️")

try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception as e:
    st.error("Erro de conexão. Verifique se o arquivo .streamlit/secrets.toml existe e está correto.")
    st.stop()

def carregar_dados_cloud():
    try:
        # TTL=0 garante dados frescos a cada recarga
        df_config = conn.read(worksheet="Config", ttl=0)
        df_membros = conn.read(worksheet="Membros", ttl=0)
        df_anos = conn.read(worksheet="Anos", ttl=0)
    except Exception:
        st.error("Erro ao ler abas. Verifique se 'Config', 'Membros' e 'Anos' existem na planilha.")
        st.stop()
    
    config_dict = dict(zip(df_config['Chave'], df_config['Valor']))
    try:
        config_dict['ultima_ata'] = int(config_dict.get('ultima_ata', 0))
    except:
        config_dict['ultima_ata'] = 0

    return {
        "config": config_dict,
        "membros": df_membros['Nome'].dropna().tolist(),
        "anos": df_anos['Ano'].dropna().tolist()
    }

def atualizar_config_cloud(chave, valor):
    df = conn.read(worksheet="Config", ttl=0)
    if chave in df['Chave'].values:
        df.loc[df['Chave'] == chave, 'Valor'] = str(valor)
    else:
        new_row = pd.DataFrame([{'Chave': chave, 'Valor': str(valor)}])
        df = pd.concat([df, new_row], ignore_index=True)
    conn.update(worksheet="Config", data=df)
    st.cache_data.clear()

def gerenciar_lista_cloud(aba, coluna, valor, acao="adicionar"):
    df = conn.read(worksheet=aba, ttl=0)
    if acao == "adicionar":
        if valor not in df[coluna].values:
            new_row = pd.DataFrame([{coluna: valor}])
            df = pd.concat([df, new_row], ignore_index=True)
            conn.update(worksheet=aba, data=df)
            return True
    elif acao == "remover":
        df = df[df[coluna] != valor]
        conn.update(worksheet=aba, data=df)
        return True
    st.cache_data.clear()
    return False

def salvar_historico_cloud(dados):
    try:
        df_hist = conn.read(worksheet="Historico", ttl=0)
        nova_linha = pd.DataFrame([{
            "Numero": dados['num_ata'],
            "Data": dados['data_reuniao'],
            "Presidente": dados['pres_nome'],
            "Secretario": dados['secretario_nome'],
            "Leitura": dados['leitura_fonte'],
            "Presentes": dados['lista_presentes_txt'],
            "Ausencias": dados['ausencias'], # Nova coluna recomendada
            "Visitantes": dados['lista_visitantes_txt'],
            "Receita": dados['receita'],
            "Despesa": dados['despesa'],
            "Saldo": dados['saldo'],
            "Socioeconomico": dados['socioeconomico'],
            "Noticias": dados['noticias_trabalhos'],
            "Palavra_Franca": dados['palavra_franca']
        }])
        df_atualizado = pd.concat([df_hist, nova_linha], ignore_index=True)
        conn.update(worksheet="Historico", data=df_atualizado)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Erro ao salvar no histórico: {e}")
        return False

# ==============================================================================
# 2. LÓGICA DE DATAS AUTOMÁTICAS
# ==============================================================================
def obter_proxima_data(dia_semana_alvo):
    if dia_semana_alvo is None or dia_semana_alvo == "":
        return datetime.now().date()
    
    try:
        dia_semana_alvo = int(dia_semana_alvo)
    except:
        return datetime.now().date()

    hoje = datetime.now().date()
    dia_hoje = hoje.weekday()
    
    if dia_hoje == dia_semana_alvo:
        return hoje
    
    dias_para_adicionar = (dia_semana_alvo - dia_hoje + 7) % 7
    return hoje + timedelta(days=dias_para_adicionar)

# ==============================================================================
# 3. FUNÇÕES AUXILIARES E GERADORES
# ==============================================================================
def formatar_valor_extenso(valor):
    try:
        extenso = num2words(valor, lang='pt_BR', to='currency')
        return f"R$ {valor:,.2f} ({extenso})".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "R$ 0,00 (zero reais)"

def formatar_data_br(data_input):
    if isinstance(data_input, (datetime, date)):
        return data_input.strftime('%d/%m/%Y')
    try:
        return datetime.strptime(str(data_input), '%Y-%m-%d').strftime('%d/%m/%Y')
    except:
        return str(data_input)

# --- Gerador DOCX ---
def gerar_docx(dados):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(12)
    titulo = doc.add_paragraph(f"Ata nº {dados['num_ata']}")
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Ata nº {dados['num_ata']} da reunião ordinária da Conferência {dados['conf_nome']} da SSVP, fundada em {dados['data_fundacao']}, agregada em {dados['data_agregacao']}, vinculada ao Conselho Particular {dados['cons_particular']}, área do Central de {dados['cons_central']}, realizada às {dados['hora_inicio']} do dia {dados['data_reuniao']} do Ano Temático: {dados['ano_tematico']}, na sala de reuniões {dados['local']}.")
    doc.add_paragraph(f"Louvado seja nosso Senhor Jesus Cristo! A reunião foi iniciada pelo Presidente, {dados['pres_nome']}, com as orações regulamentares da Sociedade de São Vicente de Paulo-SSVP.")
    doc.add_paragraph(f"A leitura espiritual foi tirada do(a) {dados['leitura_fonte']}, proclamada pelo(a) Cfd/Csc. {dados['leitor_nome']}, sendo refletida por alguns membros.")
    doc.add_paragraph(f"A ata anterior foi lida e {dados['status_ata_ant']}.")
    doc.add_paragraph(f"Em seguida foi feita a chamada, com a presença dos Confrades e Consócias: {dados['lista_presentes_txt']} e a ausência justificada: {dados['ausencias']}.")
    doc.add_paragraph(f"Presenças dos visitantes: {dados['lista_visitantes_txt']}." if dados['lista_visitantes_txt'] else "Presenças dos visitantes: Não houve.")
    receita_txt = formatar_valor_extenso(dados['receita'])
    despesa_txt = formatar_valor_extenso(dados['despesa'])
    decima_txt = formatar_valor_extenso(dados['decima'])
    saldo_txt = formatar_valor_extenso(dados['saldo'])
    doc.add_paragraph(f"Movimento do Caixa: em seguida o Tesoureiro apresentou o estado do caixa: Receita total: {receita_txt}; Despesa total: {despesa_txt}; Décima semanal: {decima_txt}; Saldo final: {saldo_txt}.")
    doc.add_paragraph(f"Agradecimentos aos visitantes. Levantamento Socioeconômico: {dados['socioeconomico']}.")
    doc.add_paragraph(f"Notícias dos trabalhos da semana: {dados['noticias_trabalhos']}")
    doc.add_paragraph(f"Novas nomeações (escala de visitas): {dados['escala_visitas']}")
    doc.add_paragraph(f"Palavra franca: {dados['palavra_franca']}")
    doc.add_paragraph(f"Expediente: {dados['expediente']}")
    doc.add_paragraph(f"Palavra dos Visitantes: {dados['palavra_visitantes']}")
    doc.add_paragraph(f"Movimento financeiro (coletas e doações): {dados['mov_financeiro_extra']}")
    doc.add_paragraph(f"Coleta Secreta: em seguida o tesoureiro fez a coleta secreta, enquanto os demais cantavam {dados['musica_final']}. Nada mais havendo a tratar, a reunião foi encerrada com as orações finais regulamentares da SSVP e com a oração para Canonização do Beato Frederico Ozanam, às {dados['hora_fim']}. Para constar, eu, {dados['secretario_nome']}, {dados['secretario_cargo']}, lavrei a presente ata, que dato e assino.")
    para_direita = doc.add_paragraph(f"{dados['cidade_estado']}, {dados['data_reuniao']}.")
    para_direita.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.add_paragraph("\n\n__________________________________________________")
    doc.add_paragraph(f"{dados['secretario_nome']} (Secretário)")
    doc.add_paragraph("\n__________________________________________________")
    doc.add_paragraph(f"{dados['pres_nome']} (Presidente)")
    return doc

# --- Gerador PDF ---
class PDF(FPDF):
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}/{{nb}}', 0, 0, 'C')

def gerar_pdf_nativo(dados):
    pdf = PDF()
    pdf.alias_nb_pages()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.set_margins(25, 25, 25)
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, f"Ata nº {dados['num_ata']}", ln=True, align="C")
    pdf.ln(5)
    pdf.set_font("Arial", size=12)
    def add_paragraph(texto):
        x_inicial = pdf.get_x()
        pdf.set_x(x_inicial + 12.5)
        pdf.multi_cell(0, 7, texto, align="J")
        pdf.ln(2)
    add_paragraph(f"Ata nº {dados['num_ata']} da reunião ordinária da Conferência {dados['conf_nome']} da SSVP, fundada em {dados['data_fundacao']}, agregada em {dados['data_agregacao']}, vinculada ao Conselho Particular {dados['cons_particular']}, área do Central de {dados['cons_central']}, realizada às {dados['hora_inicio']} do dia {dados['data_reuniao']} do Ano Temático: {dados['ano_tematico']}, na sala de reuniões {dados['local']}.")
    add_paragraph(f"Louvado seja nosso Senhor Jesus Cristo! A reunião foi iniciada pelo Presidente, {dados['pres_nome']}, com as orações regulamentares da Sociedade de São Vicente de Paulo-SSVP.")
    add_paragraph(f"A leitura espiritual foi tirada do(a) {dados['leitura_fonte']}, proclamada pelo(a) Cfd/Csc. {dados['leitor_nome']}, sendo refletida por alguns membros.")
    add_paragraph(f"A ata anterior foi lida e {dados['status_ata_ant']}.")
    add_paragraph(f"Em seguida foi feita a chamada, com a presença dos Confrades e Consócias: {dados['lista_presentes_txt']} e a ausência justificada: {dados['ausencias']}.")
    visitantes_txt = f"Presenças dos visitantes: {dados['lista_visitantes_txt']}." if dados['lista_visitantes_txt'] else "Presenças dos visitantes: Não houve."
    add_paragraph(visitantes_txt)
    receita_txt = formatar_valor_extenso(dados['receita'])
    despesa_txt = formatar_valor_extenso(dados['despesa'])
    decima_txt = formatar_valor_extenso(dados['decima'])
    saldo_txt = formatar_valor_extenso(dados['saldo'])
    add_paragraph(f"Movimento do Caixa: em seguida o Tesoureiro apresentou o estado do caixa: Receita total: {receita_txt}; Despesa total: {despesa_txt}; Décima semanal: {decima_txt}; Saldo final: {saldo_txt}.")
    add_paragraph(f"Agradecimentos aos visitantes. Levantamento Socioeconômico: {dados['socioeconomico']}.")
    add_paragraph(f"Notícias dos trabalhos da semana: {dados['noticias_trabalhos']}")
    add_paragraph(f"Novas nomeações (escala de visitas): {dados['escala_visitas']}")
    add_paragraph(f"Palavra franca: {dados['palavra_franca']}")
    add_paragraph(f"Expediente: {dados['expediente']}")
    add_paragraph(f"Palavra dos Visitantes: {dados['palavra_visitantes']}")
    add_paragraph(f"Movimento financeiro (coletas e doações): {dados['mov_financeiro_extra']}")
    add_paragraph(f"Coleta Secreta: em seguida o tesoureiro fez a coleta secreta, enquanto os demais cantavam {dados['musica_final']}. Nada mais havendo a tratar, a reunião foi encerrada com as orações finais regulamentares da SSVP e com a oração para Canonização do Beato Frederico Ozanam, às {dados['hora_fim']}. Para constar, eu, {dados['secretario_nome']}, {dados['secretario_cargo']}, lavrei a presente ata, que dato e assino.")
    pdf.ln(10)
    pdf.cell(0, 10, f"{dados['cidade_estado']}, {dados['data_reuniao']}.", ln=True, align="R")
    pdf.ln(15)
    pdf.cell(0, 5, "__________________________________________________", ln=True, align="L")
    pdf.cell(0, 5, f"{dados['secretario_nome']} (Secretário)", ln=True, align="L")
    pdf.ln(10)
    pdf.cell(0, 5, "__________________________________________________", ln=True, align="L")
    pdf.cell(0, 5, f"{dados['pres_nome']} (Presidente)", ln=True, align="L")
    return bytes(pdf.output(dest='S'))

# ==============================================================================
# 4. APP PRINCIPAL
# ==============================================================================
db = carregar_dados_cloud()
prox_num_ata = db['config']['ultima_ata'] + 1

# --- Cálculo dos Padrões ---
dia_semana_cfg = db['config'].get('dia_semana_reuniao', None)
data_padrao = obter_proxima_data(dia_semana_cfg)

hora_padrao_str = db['config'].get('horario_padrao', '20:00')
try:
    hora_padrao = datetime.strptime(hora_padrao_str, '%H:%M').time()
except:
    hora_padrao = time(20, 0)

local_padrao = db['config'].get('local_padrao', 'Sede da Conferência')
cidade_padrao = db['config'].get('cidade_padrao', 'Belo Horizonte - MG')

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("⚙️ Painel de Controle")
    with st.expander("🏢 Configurações Fixas"):
        st.info("Defina aqui os padrões para não digitar toda vez.")
        
        cfg_nome = st.text_input("Nome da Conferência", db['config'].get('nome_conf', ''))
        
        dias_semana = {0: "Segunda", 1: "Terça", 2: "Quarta", 3: "Quinta", 4: "Sexta", 5: "Sábado", 6: "Domingo"}
        idx_dia = int(dia_semana_cfg) if dia_semana_cfg is not None and str(dia_semana_cfg).isdigit() else 0
        cfg_dia = st.selectbox("Dia da Semana Padrão", options=list(dias_semana.keys()), format_func=lambda x: dias_semana[x], index=idx_dia)
        
        cfg_hora = st.text_input("Horário Padrão (HH:MM)", hora_padrao_str)
        cfg_local = st.text_input("Local Padrão", local_padrao)
        cfg_cidade = st.text_input("Cidade Padrão", cidade_padrao)
        
        st.divider()
        cfg_cp = st.text_input("Conselho Particular", db['config'].get('cons_particular', ''))
        cfg_cc = st.text_input("Conselho Central", db['config'].get('cons_central', ''))
        cfg_dt_fund = st.text_input("Data Fundação", db['config'].get('data_fundacao', ''))
        cfg_dt_agreg = st.text_input("Data Agregação", db['config'].get('data_agregacao', ''))
        
        if st.button("Salvar Configurações"):
            with st.spinner("Salvando..."):
                atualizar_config_cloud('nome_conf', cfg_nome)
                atualizar_config_cloud('dia_semana_reuniao', str(cfg_dia))
                atualizar_config_cloud('horario_padrao', cfg_hora)
                atualizar_config_cloud('local_padrao', cfg_local)
                atualizar_config_cloud('cidade_padrao', cfg_cidade)
                atualizar_config_cloud('cons_particular', cfg_cp)
                atualizar_config_cloud('cons_central', cfg_cc)
                atualizar_config_cloud('data_fundacao', cfg_dt_fund)
                atualizar_config_cloud('data_agregacao', cfg_dt_agreg)
            st.success("Configurações atualizadas!")
            st.rerun()

    with st.expander("👥 Membros"):
        novo_membro = st.text_input("Novo Membro")
        if st.button("Adicionar"):
            if gerenciar_lista_cloud("Membros", "Nome", novo_membro, "adicionar"):
                st.rerun()
        mem_remove = st.selectbox("Remover", ["Selecione..."] + db['membros'])
        if st.button("Remover"):
            if mem_remove != "Selecione...":
                gerenciar_lista_cloud("Membros", "Nome", mem_remove, "remover")
                st.rerun()

    with st.expander("📅 Anos Temáticos"):
        novo_ano = st.text_input("Novo Ano")
        if st.button("Add Ano"):
            gerenciar_lista_cloud("Anos", "Ano", novo_ano, "adicionar")
            st.rerun()
            
    st.divider()
    nova_contagem = st.number_input("Correção Contador", value=db['config']['ultima_ata'])
    if st.button("Forçar Correção"):
        atualizar_config_cloud('ultima_ata', nova_contagem)
        st.rerun()

# --- INTERFACE PRINCIPAL ---
st.title("Gerador de Ata SSVP ✝️")
st.caption("Conectado ao Arquivo Digital")

# SEÇÃO 1: Identificação (Interativa)
col1, col2, col3 = st.columns(3)
num_ata = col1.number_input("Número da Ata", value=prox_num_ata, step=1)
ano_tematico = col2.selectbox("Ano Temático", db['anos'])
data_reuniao = col3.date_input("Data da Reunião", data_padrao, format="DD/MM/YYYY")

with st.expander(f"📍 Detalhes: {hora_padrao_str} - {local_padrao} (Clique para alterar)", expanded=False):
    c_loc1, c_loc2, c_loc3 = st.columns(3)
    hora_inicio = c_loc1.time_input("Horário Início", hora_padrao)
    local = c_loc2.text_input("Local", local_padrao)
    cidade_estado = c_loc3.text_input("Cidade/UF", cidade_padrao)

st.divider()

# SEÇÃO 2: Chamada e Frequência (AGORA INTELIGENTE!)
# Esta seção está FORA do formulário para atualizar instantaneamente
st.subheader("Chamada e Frequência")
st.caption("Desmarque quem faltou. O sistema pedirá a justificativa automaticamente.")

col_pres, col_aus = st.columns([2, 1])

with col_pres:
    # Por padrão, todos vêm marcados como Presentes
    presentes = st.multiselect(
        "Membros Presentes", 
        db['membros'], 
        default=db['membros']
    )

# Calcula quem faltou (Diferença entre a lista completa e os presentes)
ausentes = [m for m in db['membros'] if m not in presentes]
motivos_ausencia = {}

with col_aus:
    if ausentes:
        st.markdown("**🛑 Ausências Detectadas:**")
        # Pergunta quem justificou
        justificaram = st.multiselect(
            "Quem justificou?", 
            ausentes,
            placeholder="Selecione..."
        )
        
        # Se alguém justificou, abre o campo para digitar o motivo
        if justificaram:
            for membro in justificaram:
                motivos_ausencia[membro] = st.text_input(f"Motivo: {membro}", placeholder="Ex: Trabalho, Doença", key=f"mot_{membro}")
    else:
        st.success("Todos presentes! 🎉")

st.divider()

# SEÇÃO 3: Formulário para Textos Longos e Envio
# Usamos st.form aqui para não recarregar a página enquanto você digita os textos longos
with st.form("form_ata_conteudo"):
    
    c_esp1, c_esp2, c_esp3 = st.columns(3)
    pres_nome = c_esp1.selectbox("Presidente", db['membros'])
    leitura_fonte = c_esp2.text_input("Fonte Leitura")
    leitor_nome = c_esp3.selectbox("Leitor", db['membros'])
    
    st.divider()
    status_ata_ant = st.radio("Ata Anterior", ["Aprovada sem ressalvas", "Aprovada com ressalvas"], horizontal=True)
    
    # Visitantes (Separado da chamada de membros)
    visitantes = st.text_area("Visitantes (Nomes)", placeholder="Se houver visitantes, digite aqui...")
    
    st.divider()
    st.markdown("### Tesouraria")
    c_fin1, c_fin2, c_fin3, c_fin4 = st.columns(4)
    receita = c_fin1.number_input("Receita", 0.0, step=0.1)
    despesa = c_fin2.number_input("Despesa", 0.0, step=0.1)
    decima = c_fin3.number_input("Décima", 0.0, step=0.1)
    saldo = c_fin4.number_input("Saldo", 0.0, step=0.1)
    
    st.divider()
    st.markdown("### Relatórios")
    socioeconomico = st.text_area("Socioeconômico", height=100)
    noticias = st.text_area("Notícias / Visitas", height=100)
    escala = st.text_area("Escala Próxima Semana")
    palavra = st.text_area("Palavra Franca")
    expediente = st.text_area("Expediente")
    
    st.divider()
    col_enc1, col_enc2 = st.columns(2)
    p_vis = col_enc1.text_input("Palavra Visitantes", "Nada a declarar")
    mov_extra = col_enc2.text_input("Mov. Extra", "Coleta regular")
    col_enc3, col_enc4 = st.columns(2)
    musica = col_enc3.text_input("Música", "Hino de Ozanam")
    hora_fim = col_enc4.time_input("Fim")
    c_sec1, c_sec2 = st.columns(2)
    sec_nome = c_sec1.selectbox("Secretário", db['membros'])
    sec_cargo = c_sec2.text_input("Cargo", "1º Secretário(a)")
    
    submit = st.form_submit_button("💾 Gerar Ata, Salvar Histórico e Baixar")

if submit:
    # Processa o texto das ausências
    lista_texto_ausencias = []
    if not ausentes:
        texto_ausencias = "Não houve."
    else:
        for m in ausentes:
            if m in motivos_ausencia and motivos_ausencia[m]:
                # Se tem motivo, coloca entre parênteses
                lista_texto_ausencias.append(f"{m} ({motivos_ausencia[m]})")
            elif m in motivos_ausencia:
                 # Se marcou que justificou mas não escreveu nada
                lista_texto_ausencias.append(f"{m} (Justificado)")
            else:
                # Não justificado
                lista_texto_ausencias.append(m)
        texto_ausencias = ", ".join(lista_texto_ausencias)

    # 1. Dados
    dados = {
        'num_ata': str(num_ata),
        'conf_nome': db['config'].get('nome_conf', ''),
        'cons_particular': db['config'].get('cons_particular', ''),
        'cons_central': db['config'].get('cons_central', ''),
        'data_fundacao': formatar_data_br(db['config'].get('data_fundacao', '')),
        'data_agregacao': formatar_data_br(db['config'].get('data_agregacao', '')),
        'ano_tematico': ano_tematico,
        'data_reuniao': formatar_data_br(data_reuniao),
        'hora_inicio': hora_inicio.strftime('%H:%M'),
        'local': local, 'pres_nome': pres_nome,
        'leitura_fonte': leitura_fonte, 'leitor_nome': leitor_nome,
        'status_ata_ant': status_ata_ant,
        'lista_presentes_txt': ", ".join(presentes),
        'ausencias': texto_ausencias, # Campo processado automaticamente
        'lista_visitantes_txt': visitantes.replace("\n", ", ") if visitantes else "",
        'receita': receita, 'despesa': despesa, 'decima': decima, 'saldo': saldo,
        'socioeconomico': socioeconomico, 'noticias_trabalhos': noticias,
        'escala_visitas': escala, 'palavra_franca': palavra,
        'expediente': expediente, 'palavra_visitantes': p_vis,
        'mov_financeiro_extra': mov_extra, 'musica_final': musica,
        'hora_fim': hora_fim.strftime('%H:%M'),
        'secretario_nome': sec_nome, 'secretario_cargo': sec_cargo,
        'cidade_estado': cidade_estado
    }
    
    # 2. Salva no Histórico do Sheets
    with st.spinner("Arquivando ata na nuvem..."):
        if salvar_historico_cloud(dados):
            st.toast("✅ Ata salva no Histórico com sucesso!")
        
    # 3. Atualiza Contador
    if num_ata > db['config']['ultima_ata']:
        atualizar_config_cloud('ultima_ata', int(num_ata))
    
    # 4. Gera Arquivos
    doc = gerar_docx(dados)
    bio_docx = io.BytesIO()
    doc.save(bio_docx)
    pdf_bytes = gerar_pdf_nativo(dados)
    
    st.success(f"Ata nº {num_ata} gerada e arquivada!")
    
    # 5. WhatsApp
    texto_zap = f"*Ata nº {num_ata} - SSVP* ✝️\n📅 {formatar_data_br(data_reuniao)}\n💰 Coleta: R$ {receita:.2f}\n🚫 Ausências: {texto_ausencias}"
    link_zap = f"https://api.whatsapp.com/send?text={urllib.parse.quote(texto_zap)}"
    st.link_button("📲 Enviar Resumo no WhatsApp", link_zap)
    
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        st.download_button("📄 Baixar PDF", pdf_bytes, f"Ata_{num_ata}.pdf", "application/pdf", type="primary", use_container_width=True)
    with col_d2:
        st.download_button("📝 Baixar Word", bio_docx.getvalue(), f"Ata_{num_ata}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)