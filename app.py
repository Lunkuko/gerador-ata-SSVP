import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF
from num2words import num2words
from datetime import datetime, date
import io

# ==============================================================================
# 1. CONFIGURAÇÃO E CONEXÃO
# ==============================================================================
st.set_page_config(page_title="Gerador de Ata SSVP (Cloud)", layout="wide", page_icon="✝️")

# Conexão com Google Sheets
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception as e:
    st.error("Erro de conexão. Verifique se o arquivo .streamlit/secrets.toml existe.")
    st.stop()

def carregar_dados_cloud():
    """Lê configurações, membros e anos temáticos da nuvem."""
    # TTL=0 evita cache antigo
    try:
        df_config = conn.read(worksheet="Config", ttl=0)
        df_membros = conn.read(worksheet="Membros", ttl=0)
        df_anos = conn.read(worksheet="Anos", ttl=0)
    except Exception:
        st.error("Erro ao ler abas da planilha. Verifique se as abas 'Config', 'Membros' e 'Anos' existem.")
        st.stop()
    
    # Processa Configuração
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
    """Atualiza um valor específico na aba Config."""
    df = conn.read(worksheet="Config", ttl=0)
    if chave in df['Chave'].values:
        df.loc[df['Chave'] == chave, 'Valor'] = str(valor)
    else:
        new_row = pd.DataFrame([{'Chave': chave, 'Valor': str(valor)}])
        df = pd.concat([df, new_row], ignore_index=True)
    conn.update(worksheet="Config", data=df)
    st.cache_data.clear()

def gerenciar_lista_cloud(aba, coluna, valor, acao="adicionar"):
    """Adiciona ou remove itens das listas (Membros/Anos)."""
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

# ==============================================================================
# 2. FUNÇÕES AUXILIARES DE FORMATAÇÃO
# ==============================================================================
def formatar_valor_extenso(valor):
    """Retorna 'R$ 10,00 (dez reais)' """
    try:
        extenso = num2words(valor, lang='pt_BR', to='currency')
        return f"R$ {valor:,.2f} ({extenso})".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "R$ 0,00 (zero reais)"

def formatar_data_br(data_input):
    """Garante formato DD/MM/AAAA"""
    if isinstance(data_input, (datetime, date)):
        return data_input.strftime('%d/%m/%Y')
    try:
        return datetime.strptime(str(data_input), '%Y-%m-%d').strftime('%d/%m/%Y')
    except:
        return str(data_input)

# ==============================================================================
# 3. GERADORES DE DOCUMENTO (DOCX E PDF)
# ==============================================================================

# --- GERADOR DOCX ---
def gerar_docx(dados):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(12)

    # Título
    titulo = doc.add_paragraph(f"Ata nº {dados['num_ata']}")
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Texto Padrão
    # [cite: 2] Cabeçalho
    doc.add_paragraph(
        f"Ata nº {dados['num_ata']} da reunião ordinária da Conferência {dados['conf_nome']} da SSVP, "
        f"fundada em {dados['data_fundacao']}, agregada em {dados['data_agregacao']}, "
        f"vinculada ao Conselho Particular {dados['cons_particular']}, área do Central de {dados['cons_central']}, "
        f"realizada às {dados['hora_inicio']} do dia {dados['data_reuniao']} "
        f"do Ano Temático: {dados['ano_tematico']}, na sala de reuniões {dados['local']}."
    )
    # [cite: 3] Orações Iniciais
    doc.add_paragraph(f"Louvado seja nosso Senhor Jesus Cristo! A reunião foi iniciada pelo Presidente, {dados['pres_nome']}, com as orações regulamentares da Sociedade de São Vicente de Paulo-SSVP.")
    
    # [cite: 4, 5] Leitura
    doc.add_paragraph(f"A leitura espiritual foi tirada do(a) {dados['leitura_fonte']}, proclamada pelo(a) Cfd/Csc. {dados['leitor_nome']}, sendo refletida por alguns membros.")
    
    # [cite: 6] Ata Anterior
    doc.add_paragraph(f"A ata anterior foi lida e {dados['status_ata_ant']}.")
    
    # [cite: 7] Chamada
    doc.add_paragraph(f"Em seguida foi feita a chamada, com a presença dos Confrades e Consócias: {dados['lista_presentes_txt']} e a ausência justificada: {dados['ausencias']}.")
    
    # [cite: 8] Visitantes
    doc.add_paragraph(f"Presenças dos visitantes: {dados['lista_visitantes_txt']}." if dados['lista_visitantes_txt'] else "Presenças dos visitantes: Não houve.")
    
    # [cite: 9, 10] Tesouraria
    receita_txt = formatar_valor_extenso(dados['receita'])
    despesa_txt = formatar_valor_extenso(dados['despesa'])
    decima_txt = formatar_valor_extenso(dados['decima'])
    saldo_txt = formatar_valor_extenso(dados['saldo'])
    doc.add_paragraph(f"Movimento do Caixa: em seguida o Tesoureiro apresentou o estado do caixa: Receita total: {receita_txt}; Despesa total: {despesa_txt}; Décima semanal: {decima_txt}; Saldo final: {saldo_txt}.")
    
    # [cite: 11-14] Relatórios
    doc.add_paragraph(f"Agradecimentos aos visitantes. Levantamento Socioeconômico: {dados['socioeconomico']}.")
    doc.add_paragraph(f"Notícias dos trabalhos da semana: {dados['noticias_trabalhos']}")
    doc.add_paragraph(f"Novas nomeações (escala de visitas): {dados['escala_visitas']}")
    doc.add_paragraph(f"Palavra franca: {dados['palavra_franca']}")
    doc.add_paragraph(f"Expediente: {dados['expediente']}")
    
    # [cite: 15] Encerramento
    doc.add_paragraph(f"Palavra dos Visitantes: {dados['palavra_visitantes']}")
    doc.add_paragraph(f"Movimento financeiro (coletas e doações): {dados['mov_financeiro_extra']}")
    
    # [cite: 16, 17] Oração Final
    doc.add_paragraph(f"Coleta Secreta: em seguida o tesoureiro fez a coleta secreta, enquanto os demais cantavam {dados['musica_final']}. Nada mais havendo a tratar, a reunião foi encerrada com as orações finais regulamentares da SSVP e com a oração para Canonização do Beato Frederico Ozanam, às {dados['hora_fim']}. Para constar, eu, {dados['secretario_nome']}, {dados['secretario_cargo']}, lavrei a presente ata, que dato e assino.")
    
    # [cite: 18, 23, 28] Assinaturas
    para_direita = doc.add_paragraph(f"{dados['cidade_estado']}, {dados['data_reuniao']}.")
    para_direita.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.add_paragraph("\n\n__________________________________________________")
    doc.add_paragraph(f"{dados['secretario_nome']} (Secretário)")
    doc.add_paragraph("\n__________________________________________________")
    doc.add_paragraph(f"{dados['pres_nome']} (Presidente)")
    
    return doc

# --- GERADOR PDF ---
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
    pdf.set_margins(25, 25, 25) # Margens padrão
    
    # Título
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, f"Ata nº {dados['num_ata']}", ln=True, align="C")
    pdf.ln(5)
    
    pdf.set_font("Arial", size=12)

    def add_paragraph(texto):
        x_inicial = pdf.get_x()
        pdf.set_x(x_inicial + 12.5) # Recuo de parágrafo
        pdf.multi_cell(0, 7, texto, align="J") # Justificado
        pdf.ln(2)

    # Texto (Mesma estrutura do DOCX)
    add_paragraph(f"Ata nº {dados['num_ata']} da reunião ordinária da Conferência {dados['conf_nome']} da SSVP, fundada em {dados['data_fundacao']}, agregada em {dados['data_agregacao']}, vinculada ao Conselho Particular {dados['cons_particular']}, área do Central de {dados['cons_central']}, realizada às {dados['hora_inicio']} do dia {dados['data_reuniao']} do Ano Temático: {dados['ano_tematico']}, na sala de reuniões {dados['local']}.")
    
    add_paragraph(f"Louvado seja nosso Senhor Jesus Cristo! A reunião foi iniciada pelo Presidente, {dados['pres_nome']}, com as orações regulamentares da Sociedade de São Vicente de Paulo-SSVP.")
    
    add_paragraph(f"A leitura espiritual foi tirada do(a) {dados['leitura_fonte']}, proclamada pelo(a) Cfd/Csc. {dados['leitor_nome']}, sendo refletida por alguns membros.")
    
    add_paragraph(f"A ata anterior foi lida e {dados['status_ata_ant']}.")
    
    add_paragraph(f"Em seguida foi feita a chamada, com a presença dos Confrades e Consócias: {dados['lista_presentes_txt']} e a ausência justificada: {dados['ausencias']}.")
    
    visitantes_txt = f"Presenças dos visitantes: {dados['lista_visitantes_txt']}." if dados['lista_visitantes_txt'] else "Presenças dos visitantes: Não houve."
    add_paragraph(visitantes_txt)
    
    # Financeiro Formatado
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
# 4. INTERFACE DO USUÁRIO (FRONTEND)
# ==============================================================================

# Carregamento Inicial
db = carregar_dados_cloud()
prox_num_ata = db['config']['ultima_ata'] + 1

# --- BARRA LATERAL (Gestão) ---
with st.sidebar:
    st.header("⚙️ Painel de Controle")
    
    # Configurações da Conferência
    with st.expander("🏢 Dados Fixos da Conferência"):
        cfg_nome = st.text_input("Nome", db['config'].get('nome_conf', ''))
        cfg_cp = st.text_input("Conselho Particular", db['config'].get('cons_particular', ''))
        cfg_cc = st.text_input("Conselho Central", db['config'].get('cons_central', ''))
        cfg_dt_fund = st.text_input("Data Fundação (DD/MM/AAAA)", db['config'].get('data_fundacao', ''))
        cfg_dt_agreg = st.text_input("Data Agregação (DD/MM/AAAA)", db['config'].get('data_agregacao', ''))
        
        if st.button("Salvar Configurações"):
            with st.spinner("Salvando na nuvem..."):
                atualizar_config_cloud('nome_conf', cfg_nome)
                atualizar_config_cloud('cons_particular', cfg_cp)
                atualizar_config_cloud('cons_central', cfg_cc)
                atualizar_config_cloud('data_fundacao', cfg_dt_fund)
                atualizar_config_cloud('data_agregacao', cfg_dt_agreg)
            st.success("Salvo!")
            st.rerun()

    # Gestão de Membros
    with st.expander("👥 Gerenciar Membros"):
        novo_membro = st.text_input("Novo Membro (Ex: Cfd. João)")
        if st.button("Adicionar Membro"):
            if gerenciar_lista_cloud("Membros", "Nome", novo_membro, "adicionar"):
                st.success("Adicionado!")
                st.rerun()
        
        mem_remove = st.selectbox("Remover", ["Selecione..."] + db['membros'])
        if st.button("Remover Membro"):
            if mem_remove != "Selecione...":
                gerenciar_lista_cloud("Membros", "Nome", mem_remove, "remover")
                st.rerun()

    # Gestão de Anos
    with st.expander("📅 Anos Temáticos"):
        novo_ano = st.text_input("Novo Ano Temático")
        if st.button("Adicionar Ano"):
            gerenciar_lista_cloud("Anos", "Ano", novo_ano, "adicionar")
            st.rerun()

    st.divider()
    st.caption(f"Última Ata Registrada no Sistema: {db['config']['ultima_ata']}")
    nova_contagem = st.number_input("Correção Manual do Contador", value=db['config']['ultima_ata'])
    if st.button("Forçar Correção"):
        atualizar_config_cloud('ultima_ata', nova_contagem)
        st.success("Contador corrigido.")
        st.rerun()

# --- FORMULÁRIO PRINCIPAL ---
st.title("Gerador de Ata SSVP ✝️")
st.caption("Sistema Conectado à Nuvem (Google Sheets)")

with st.form("form_ata"):
    
    # [cite: 2] Identificação
    col1, col2, col3 = st.columns(3)
    num_ata = col1.number_input("Número da Ata", value=prox_num_ata, step=1)
    ano_tematico = col2.selectbox("Ano Temático", db['anos'])
    data_reuniao = col3.date_input("Data da Reunião", datetime.now())
    
    col4, col5, col6 = st.columns(3)
    hora_inicio = col4.time_input("Horário Início", datetime.now().time())
    local = col5.text_input("Local", "Sede da Conferência")
    cidade_estado = col6.text_input("Cidade/UF (Para Assinatura)", "Belo Horizonte - MG")
    
    st.divider()
    
    # [cite: 3, 4, 5] Espiritualidade
    st.subheader("Espiritualidade")
    c_esp1, c_esp2, c_esp3 = st.columns(3)
    pres_nome = c_esp1.selectbox("Presidente", db['membros'])
    leitura_fonte = c_esp2.text_input("Fonte da Leitura", placeholder="Ex: Lucas 10, 25-37")
    leitor_nome = c_esp3.selectbox("Leitor", db['membros'])
    
    st.divider()
    
    # [cite: 6, 7, 8] Frequência
    st.subheader("Frequência")
    status_ata_ant = st.radio("Ata Anterior", ["Aprovada sem ressalvas", "Aprovada com ressalvas"], horizontal=True)
    presentes = st.multiselect("Membros Presentes", db['membros'], default=db['membros'])
    ausencias = st.text_input("Ausências Justificadas", placeholder="Ex: Cfd. José (Doença)")
    visitantes = st.text_area("Visitantes (Nome completo)", placeholder="Ex: Sr. Antônio, Pe. Marcelo...")
    
    st.divider()
    
    # [cite: 9, 10] Tesouraria
    st.subheader("Tesouraria")
    c_fin1, c_fin2, c_fin3, c_fin4 = st.columns(4)
    receita = c_fin1.number_input("Receita Total (R$)", 0.0, step=0.1)
    despesa = c_fin2.number_input("Despesa Total (R$)", 0.0, step=0.1)
    decima = c_fin3.number_input("Décima (R$)", 0.0, step=0.1)
    saldo = c_fin4.number_input("Saldo Final (R$)", 0.0, step=0.1)
    
    st.divider()
    
    # [cite: 11-14] Desenvolvimento
    st.subheader("Desenvolvimento")
    socioeconomico = st.text_area("Levantamento Socioeconômico", height=100)
    noticias = st.text_area("Notícias dos Trabalhos / Visitas", height=100)
    escala = st.text_area("Escala Próxima Semana", height=68)
    palavra = st.text_area("Palavra Franca", height=68)
    expediente = st.text_area("Expediente (Correspondências)", height=68)
    
    st.divider()
    
    # [cite: 15-17] Encerramento
    st.subheader("Encerramento")
    col_enc1, col_enc2 = st.columns(2)
    p_vis = col_enc1.text_input("Palavra dos Visitantes", "Nada a declarar")
    mov_extra = col_enc2.text_input("Movimento Extra (Coleta)", "Coleta regular realizada")
    
    col_enc3, col_enc4 = st.columns(2)
    musica = col_enc3.text_input("Música Final", "Hino de Ozanam")
    hora_fim = col_enc4.time_input("Horário Fim")
    
    # [cite: 17] Secretária
    c_sec1, c_sec2 = st.columns(2)
    sec_nome = c_sec1.selectbox("Secretário(a)", db['membros'])
    sec_cargo = c_sec2.text_input("Cargo Secretária", "1º Secretário(a)")
    
    submit = st.form_submit_button("💾 Gerar Ata (Word e PDF)")

if submit:
    # 1. Atualiza Contador na Nuvem
    if num_ata > db['config']['ultima_ata']:
        atualizar_config_cloud('ultima_ata', int(num_ata))
        st.toast("Contador atualizado na nuvem!")

    # 2. Prepara Dados
    dados = {
        'num_ata': str(num_ata),
        'conf_nome': db['config'].get('nome_conf', 'Nome da Conferência'),
        'cons_particular': db['config'].get('cons_particular', 'CP'),
        'cons_central': db['config'].get('cons_central', 'CC'),
        'data_fundacao': formatar_data_br(db['config'].get('data_fundacao', '')),
        'data_agregacao': formatar_data_br(db['config'].get('data_agregacao', '')),
        'ano_tematico': ano_tematico,
        'data_reuniao': formatar_data_br(data_reuniao),
        'hora_inicio': hora_inicio.strftime('%H:%M'),
        'local': local, 'pres_nome': pres_nome,
        'leitura_fonte': leitura_fonte, 'leitor_nome': leitor_nome,
        'status_ata_ant': status_ata_ant,
        'lista_presentes_txt': ", ".join(presentes),
        'ausencias': ausencias,
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
    
    # 3. Gera Arquivos
    # Word
    doc = gerar_docx(dados)
    bio_docx = io.BytesIO()
    doc.save(bio_docx)
    
    # PDF
    try:
        pdf_bytes = gerar_pdf_nativo(dados)
        pdf_success = True
    except Exception as e:
        st.error(f"Erro ao gerar PDF: {e}")
        pdf_success = False
    
    st.success(f"✅ Ata nº {num_ata} gerada com sucesso!")
    
    # 4. Botões de Download
    col_d1, col_d2 = st.columns(2)
    
    if pdf_success:
        with col_d1:
            st.download_button(
                label="📄 Baixar PDF (Pronto p/ Imprimir)",
                data=pdf_bytes,
                file_name=f"Ata_{num_ata}.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True
            )
            
    with col_d2:
        st.download_button(
            label="📝 Baixar Word (Para Editar)",
            data=bio_docx.getvalue(),
            file_name=f"Ata_{num_ata}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )