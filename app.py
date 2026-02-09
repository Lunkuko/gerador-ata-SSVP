import streamlit as st
import streamlit_authenticator as stauth
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF
from num2words import num2words
from datetime import datetime, date, timedelta, time
import io
import time

# ==============================================================================
# 1. CONFIGURAÇÃO INICIAL E CLASSES GLOBAIS
# ==============================================================================
st.set_page_config(page_title="Gerador de Ata SSVP (Seguro)", layout="wide", page_icon="✝️")

# --- CLASSE PDF (Escopo Global para evitar erros) ---
class PDF(FPDF):
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}/{{nb}}', 0, 0, 'C')

try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception as e:
    st.error("Erro de conexão. Verifique se o arquivo .streamlit/secrets.toml existe.")
    st.stop()

# ==============================================================================
# 2. FUNÇÕES UTILITÁRIAS E DE VALIDAÇÃO
# ==============================================================================

def eh_valido(valor):
    """
    Retorna False se o valor for vazio, nulo, zero, nan ou NaT.
    Usado para impedir que campos vazios apareçam na impressão.
    """
    if valor is None: return False
    s = str(valor).strip().lower()
    valores_invalidos = ["", "nan", "none", "null", "false", "0", "0.0", "nat", "nattype"]
    return s not in valores_invalidos

def formatar_valor_extenso(valor):
    try:
        if not valor: return "R$ 0,00"
        valor = float(valor)
        extenso = num2words(valor, lang='pt_BR', to='currency')
        return f"R$ {valor:,.2f} ({extenso})".replace(",", "X").replace(".", ",").replace("X", ".")
    except: return "R$ 0,00"

def formatar_data_br(data):
    if isinstance(data, (datetime, date)): return data.strftime('%d/%m/%Y')
    return str(data)

def obter_proxima_data(dia_alvo):
    if dia_alvo is None or dia_alvo == "": return datetime.now().date()
    try: dia_alvo = int(dia_alvo)
    except: return datetime.now().date()
    hoje = datetime.now().date()
    dia_hoje = hoje.weekday()
    if dia_hoje == dia_alvo: return hoje
    dias_add = (dia_alvo - dia_hoje + 7) % 7
    return hoje + timedelta(days=dias_add)

def get_index_membro(nome, lista):
    try: return lista.index(nome) if nome in lista else 0
    except: return 0

def limpar_memoria():
    carregar_dados_cloud.clear()
    obter_saldo_anterior.clear()
    st.cache_data.clear()
    if 'saldo_cache' in st.session_state:
        del st.session_state['saldo_cache']

# ==============================================================================
# 3. FUNÇÕES DE BANCO DE DADOS (GOOGLE SHEETS)
# ==============================================================================

def carregar_usuarios():
    """Lê os usuários da planilha 'Usuarios'."""
    try:
        df = conn.read(worksheet="Usuarios", ttl=0)
        df = df.dropna(subset=['username']) 
        credentials = {"usernames": {}}
        for _, row in df.iterrows():
            credentials["usernames"][row['username']] = {
                "name": row['name'],
                "password": row['password'], 
                "roles": [row['role']] if 'role' in row else ['editor']
            }
        return credentials
    except:
        return {"usernames": {}}

def salvar_novo_usuario(username, name, password_hash, role):
    try:
        df = conn.read(worksheet="Usuarios", ttl=0)
        if not df.empty and username in df['username'].values:
            return False, "Usuário já existe!"
        novo_user = pd.DataFrame([{
            "username": username, "name": name,
            "password": password_hash, "role": role
        }])
        df_atualizado = pd.concat([df, novo_user], ignore_index=True)
        conn.update(worksheet="Usuarios", data=df_atualizado)
        return True, "Usuário criado com sucesso!"
    except Exception as e:
        return False, f"Erro ao salvar: {e}"

@st.cache_data(ttl=3600)
def carregar_dados_cloud():
    tentativas = 0
    max_tentativas = 3
    while tentativas < max_tentativas:
        try:
            df_config = conn.read(worksheet="Config")
            df_membros = conn.read(worksheet="Membros")
            df_anos = conn.read(worksheet="Anos")
            break 
        except Exception as e:
            if "429" in str(e) or "Quota" in str(e):
                tentativas += 1
                time.sleep(2 ** tentativas)
                if tentativas == max_tentativas: st.stop()
            else: st.stop()

    if df_membros.empty: lista_membros = []
    else: lista_membros = df_membros['Nome'].dropna().astype(str).tolist()
        
    if df_anos.empty: lista_anos = []
    else: lista_anos = df_anos['Ano'].dropna().astype(str).tolist()

    config_dict = dict(zip(df_config['Chave'], df_config['Valor']))
    try: config_dict['ultima_ata'] = int(config_dict.get('ultima_ata', 0))
    except: config_dict['ultima_ata'] = 0

    return {"config": config_dict, "membros": lista_membros, "anos": lista_anos}

@st.cache_data(ttl=3600)
def obter_saldo_anterior():
    """Busca o saldo com Cache para evitar estourar a API."""
    try:
        df_hist = conn.read(worksheet="Historico") # ttl padrão usa cache
        if not df_hist.empty and 'Saldo' in df_hist.columns and len(df_hist) > 0:
            val = df_hist['Saldo'].iloc[-1]
            return float(val) if val else 0.0
    except: pass
    return 0.0

def salvar_lote_configs(dicionario_mudancas):
    """Salva configurações em lote para economizar cota."""
    try:
        df = conn.read(worksheet="Config")
        for chave, valor in dicionario_mudancas.items():
            valor_str = str(valor)
            if chave in df['Chave'].values:
                df.loc[df['Chave'] == chave, 'Valor'] = valor_str
            else:
                new_row = pd.DataFrame([{'Chave': chave, 'Valor': valor_str}])
                df = pd.concat([df, new_row], ignore_index=True)
        conn.update(worksheet="Config", data=df)
        limpar_memoria()
        return True
    except Exception as e:
        st.error(f"Erro ao salvar: {e}")
        return False

def gerenciar_lista_cloud(aba, coluna, valor, acao="adicionar"):
    time.sleep(1)
    df = conn.read(worksheet=aba)
    if acao == "adicionar":
        if valor not in df[coluna].values:
            new_row = pd.DataFrame([{coluna: valor}])
            df = pd.concat([df, new_row], ignore_index=True)
            conn.update(worksheet=aba, data=df)
    elif acao == "remover":
        df = df[df[coluna] != valor]
        conn.update(worksheet=aba, data=df)
    limpar_memoria()
    return True

def buscar_ata_para_edicao(num_ata_busca):
    try:
        df_hist = conn.read(worksheet="Historico", ttl=0)
        df_hist.columns = df_hist.columns.str.strip()
        col_num = None
        possiveis = ["Numero", "Número", "Num", "Nº"]
        for c in df_hist.columns:
            if any(x in c for x in possiveis): col_num = c; break
        
        if not col_num: return None, df_hist, "Coluna ID não encontrada."

        df_hist['Busca_ID'] = df_hist[col_num].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        termo = str(num_ata_busca).strip()
        ata = df_hist[df_hist['Busca_ID'] == termo]
        
        if not ata.empty: 
            return ata.iloc[0].to_dict(), None, "Ata encontrada!"
        else:
            return None, df_hist, f"Ata {termo} não encontrada."
    except Exception as e: return None, None, str(e)

def salvar_historico_cloud(dados):
    try:
        df_hist = conn.read(worksheet="Historico", ttl=0)
        df_hist.columns = df_hist.columns.str.strip()
        col_num = "Numero"
        for c in df_hist.columns:
            if "umero" in c or "úmero" in c or "Num" in c: col_num = c; break

        df_hist['Busca_ID'] = df_hist[col_num].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        num_atual = str(dados['num_ata']).strip()
        
        nova_linha = {
            col_num: num_atual,
            "Data": dados['data_reuniao'],
            "Presidente": dados['pres_nome'],
            "Secretario": dados['secretario_nome'],
            "Leitura": dados['leitura_fonte'],
            "Presentes": dados['lista_presentes_txt'],
            "Ausencias": dados['ausencias'],
            "Visitantes": dados['lista_visitantes_txt'],
            "Receita": dados['receita'],
            "Despesa": dados['despesa'],
            "Saldo": dados['saldo'],
            "Socioeconomico": dados['socioeconomico'],
            "Noticias": dados['noticias_trabalhos'],
            "Palavra_Franca": dados['palavra_franca']
        }

        if num_atual in df_hist['Busca_ID'].values:
            idx = df_hist.index[df_hist['Busca_ID'] == num_atual].tolist()[0]
            colunas_originais = [c for c in df_hist.columns if c != 'Busca_ID']
            for col, val in nova_linha.items():
                if col in colunas_originais: df_hist.at[idx, col] = val
            df_atualizado = df_hist[colunas_originais]
            msg_tipo = "atualizada"
        else:
            df_nova = pd.DataFrame([nova_linha])
            colunas_limpas = [c for c in df_hist.columns if c != 'Busca_ID']
            df_atualizado = pd.concat([df_hist[colunas_limpas], df_nova], ignore_index=True)
            msg_tipo = "criada"

        conn.update(worksheet="Historico", data=df_atualizado)
        return True, msg_tipo
    except Exception as e: return False, "erro"

# ==============================================================================
# 4. GERADORES DE DOCUMENTOS (COM VALIDACAO AGRESSIVA)
# ==============================================================================

def gerar_docx(dados):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(12)
    
    texto = f"Ata nº {dados['num_ata']} da reunião ordinária da Conferência {dados['conf_nome']} da SSVP"
    if eh_valido(dados['data_fundacao']): texto += f", fundada em {dados['data_fundacao']}"
    if eh_valido(dados['data_agregacao']): texto += f", agregada em {dados['data_agregacao']}"
    texto += f", vinculada ao Conselho Particular {dados['cons_particular']}, área do Central de {dados['cons_central']}, realizada às {dados['hora_inicio']} do dia {dados['data_reuniao']} do Ano Temático: {dados['ano_tematico']}, na sala de reuniões {dados['local']}."
    
    texto += f" Louvado seja nosso Senhor Jesus Cristo! A reunião foi iniciada pelo Presidente, {dados['pres_nome']}, com as orações regulamentares da Sociedade de São Vicente de Paulo-SSVP."
    texto += f" A leitura espiritual foi tirada do(a) {dados['leitura_fonte']}, proclamada pelo(a) Cfd/Csc. {dados['leitor_nome']}, sendo refletida por alguns membros."
    texto += f" A ata anterior foi lida e {dados['status_ata_ant']}."
    texto += f" Em seguida foi feita a chamada, com a presença dos Confrades e Consócias: {dados['lista_presentes_txt']}."
    
    if eh_valido(dados['lista_visitantes_txt']): texto += f" Presenças dos visitantes: {dados['lista_visitantes_txt']}."
    
    rec = formatar_valor_extenso(dados['receita'])
    des = formatar_valor_extenso(dados['despesa'])
    dec = formatar_valor_extenso(dados['decima'])
    sal = formatar_valor_extenso(dados['saldo'])
    tes = f"o(a) Tesoureiro(a) {dados['tes_nome']}" if eh_valido(dados['tes_nome']) else "o Tesoureiro"
    texto += f" Movimento do Caixa: em seguida {tes} apresentou o estado do caixa: Receita total: {rec}; Despesa total: {des}; Décima semanal: {dec}; Saldo final: {sal}."
    
    if eh_valido(dados['lista_visitantes_txt']): texto += " Agradecimentos aos visitantes."
    if eh_valido(dados['socioeconomico']): texto += f" Levantamento Socioeconômico: {dados['socioeconomico']}."
    if eh_valido(dados['noticias_trabalhos']): texto += f" Notícias dos trabalhos da semana: {dados['noticias_trabalhos']}."
    if eh_valido(dados['escala_visitas']): texto += f" Novas nomeações (escala de visitas): {dados['escala_visitas']}."
    if eh_valido(dados['palavra_franca']): texto += f" Palavra franca: {dados['palavra_franca']}."
    if eh_valido(dados['expediente']): texto += f" Expediente: {dados['expediente']}."
    if eh_valido(dados['palavra_visitantes']): texto += f" Palavra dos Visitantes: {dados['palavra_visitantes']}."
    
    tes_col = f"o(a) tesoureiro(a) {dados['tes_nome']}" if eh_valido(dados['tes_nome']) else "o tesoureiro"
    texto += f" Coleta Secreta: em seguida {tes_col} fez a coleta secreta, enquanto os demais cantavam {dados['musica_final']}."
    texto += f" Nada mais havendo a tratar, a reunião foi encerrada com as orações finais regulamentares da SSVP e com a oração para Canonização do Beato Frederico Ozanam, às {dados['hora_fim']}."
    texto += f" Para constar, eu, {dados['secretario_nome']}, {dados['secretario_cargo']}, lavrei a presente ata, que dato e assino."
    
    p = doc.add_paragraph(texto)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pd = doc.add_paragraph(f"{dados['cidade_estado']}, {dados['data_reuniao']}.")
    pd.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.add_paragraph("\n\nAssinaturas dos Presentes:")
    for _ in range(30): doc.add_paragraph("_"*85)
    return doc

def gerar_pdf_nativo(dados):
    def limpar_texto(txt):
        if not eh_valido(txt): return ""
        return str(txt).encode('latin-1', 'replace').decode('latin-1')

    pdf = PDF()
    pdf.alias_nb_pages()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.set_margins(25, 25, 25)
    
    texto = f"Ata nº {limpar_texto(dados['num_ata'])} da reunião ordinária da Conferência {limpar_texto(dados['conf_nome'])} da SSVP"
    if eh_valido(dados['data_fundacao']): texto += f", fundada em {limpar_texto(dados['data_fundacao'])}"
    if eh_valido(dados['data_agregacao']): texto += f", agregada em {limpar_texto(dados['data_agregacao'])}"
    texto += f", vinculada ao Conselho Particular {limpar_texto(dados['cons_particular'])}, área do Central de {limpar_texto(dados['cons_central'])}, realizada às {limpar_texto(dados['hora_inicio'])} do dia {limpar_texto(dados['data_reuniao'])} do Ano Temático: {limpar_texto(dados['ano_tematico'])}, na sala de reuniões {limpar_texto(dados['local'])}."
    
    texto += f" Louvado seja nosso Senhor Jesus Cristo! A reunião foi iniciada pelo Presidente, {limpar_texto(dados['pres_nome'])}, com as orações regulamentares da Sociedade de São Vicente de Paulo-SSVP."
    texto += f" A leitura espiritual foi tirada do(a) {limpar_texto(dados['leitura_fonte'])}, proclamada pelo(a) Cfd/Csc. {limpar_texto(dados['leitor_nome'])}, sendo refletida por alguns membros."
    texto += f" A ata anterior foi lida e {limpar_texto(dados['status_ata_ant'])}."
    texto += f" Em seguida foi feita a chamada, com a presença dos Confrades e Consócias: {limpar_texto(dados['lista_presentes_txt'])}."
    
    if eh_valido(dados['lista_visitantes_txt']): texto += f" Presenças dos visitantes: {limpar_texto(dados['lista_visitantes_txt'])}."
    
    rec = formatar_valor_extenso(dados['receita'])
    des = formatar_valor_extenso(dados['despesa'])
    dec = formatar_valor_extenso(dados['decima'])
    sal = formatar_valor_extenso(dados['saldo'])
    tes = f"o(a) Tesoureiro(a) {limpar_texto(dados['tes_nome'])}" if eh_valido(dados['tes_nome']) else "o Tesoureiro"
    texto += f" Movimento do Caixa: em seguida {tes} apresentou o estado do caixa: Receita total: {rec}; Despesa total: {des}; Décima semanal: {dec}; Saldo final: {sal}."
    
    if eh_valido(dados['lista_visitantes_txt']): texto += " Agradecimentos aos visitantes."
    if eh_valido(dados['socioeconomico']): texto += f" Levantamento Socioeconômico: {limpar_texto(dados['socioeconomico'])}."
    if eh_valido(dados['noticias_trabalhos']): texto += f" Notícias dos trabalhos da semana: {limpar_texto(dados['noticias_trabalhos'])}."
    if eh_valido(dados['escala_visitas']): texto += f" Novas nomeações (escala de visitas): {limpar_texto(dados['escala_visitas'])}."
    if eh_valido(dados['palavra_franca']): texto += f" Palavra franca: {limpar_texto(dados['palavra_franca'])}."
    if eh_valido(dados['expediente']): texto += f" Expediente: {limpar_texto(dados['expediente'])}."
    if eh_valido(dados['palavra_visitantes']): texto += f" Palavra dos Visitantes: {limpar_texto(dados['palavra_visitantes'])}."
    
    tes_col = f"o(a) tesoureiro(a) {limpar_texto(dados['tes_nome'])}" if eh_valido(dados['tes_nome']) else "o tesoureiro"
    texto += f" Coleta Secreta: em seguida {tes_col} fez a coleta secreta, enquanto os demais cantavam {limpar_texto(dados['musica_final'])}."
    texto += f" Nada mais havendo a tratar, a reunião foi encerrada com as orações finais regulamentares da SSVP e com a oração para Canonização do Beato Frederico Ozanam, às {limpar_texto(dados['hora_fim'])}."
    texto += f" Para constar, eu, {limpar_texto(dados['secretario_nome'])}, {limpar_texto(dados['secretario_cargo'])}, lavrei a presente ata, que dato e assino."
    
    pdf.multi_cell(0, 7, texto, align="J")
    pdf.ln(10)
    pdf.cell(0, 10, f"{limpar_texto(dados['cidade_estado'])}, {limpar_texto(dados['data_reuniao'])}.", ln=True, align="R")
    pdf.ln(10)
    pdf.cell(0, 10, "Assinaturas dos Presentes:", ln=True, align="L")
    for _ in range(30): pdf.cell(0, 8, "_"*65, ln=True, align="C")
    
    try: return pdf.output(dest='S').encode('latin-1', 'replace')
    except: return b"%PDF-1.4 erro"

# ==============================================================================
# 5. AUTENTICAÇÃO E UI
# ==============================================================================

credentials_dict = carregar_usuarios()
authenticator = stauth.Authenticate(credentials_dict, "ssvp_cookie_seguro", "chave_secreta_2026", 30)
name, authentication_status, username = authenticator.login("main")

if authentication_status:
    with st.sidebar:
        st.write(f"👤 Olá, **{name}**")
        authenticator.logout("Sair", "sidebar")
        st.divider()
        
        # Recuperação Segura de Role
        dados_usuario = credentials_dict['usernames'].get(username, {})
        roles = dados_usuario.get('roles', ['editor'])
        user_role = roles[0] if isinstance(roles, list) else roles
        
        if user_role == 'admin':
            with st.expander("🔐 Gestão de Usuários"):
                with st.form("form_novo_user"):
                    nu, nn = st.text_input("Login"), st.text_input("Nome")
                    np, nr = st.text_input("Senha", type="password"), st.selectbox("Role", ["editor", "admin"])
                    if st.form_submit_button("Criar"):
                        try: h = stauth.Hasher([np]).generate()[0]
                        except: h = stauth.Hasher().generate([np])[0]
                        ok, msg = salvar_novo_usuario(nu, nn, h, nr)
                        if ok: st.success(msg); time.sleep(1); st.rerun()
                        else: st.error(msg)

    db = carregar_dados_cloud()
    if 'dados_carregados' not in st.session_state: st.session_state.dados_carregados = {}
    dc = st.session_state.dados_carregados

    # Configs Padrão
    dia_cfg = db['config'].get('dia_semana_reuniao', None)
    data_pad = obter_proxima_data(dia_cfg)
    hora_pad = datetime.strptime(db['config'].get('horario_padrao', '20:00'), '%H:%M').time()
    
    with st.sidebar:
        st.header("⚙️ Painel")
        with st.expander("🛠️ Corrigir Ata"):
            nb = st.number_input("Nº Ata", min_value=1, step=1)
            if st.button("Carregar"):
                d_old, debug_df, msg = buscar_ata_para_edicao(nb)
                if d_old: st.session_state.dados_carregados = d_old; st.rerun()
                else: st.error(msg)

        with st.expander("👔 Cargos"):
            cp = st.selectbox("Presidente", db['membros'], index=get_index_membro(db['config'].get('pres_padrao'), db['membros']))
            st.divider()
            cs1 = st.selectbox("1º Sec.", db['membros'], index=get_index_membro(db['config'].get('sec_padrao'), db['membros']))
            csc1 = st.text_input("Cargo 1", db['config'].get('sec_cargo_padrao',''))
            st.divider()
            cs2 = st.selectbox("2º Sec.", db['membros'], index=get_index_membro(db['config'].get('sec2_padrao'), db['membros']))
            csc2 = st.text_input("Cargo 2", db['config'].get('sec2_cargo_padrao',''))
            st.divider()
            ct = st.selectbox("Tesoureiro", db['membros'], index=get_index_membro(db['config'].get('tes_padrao'), db['membros']))
            
            if st.button("Salvar Cargos"):
                mudancas = {'pres_padrao':cp, 'sec_padrao':cs1, 'sec_cargo_padrao':csc1, 'sec2_padrao':cs2, 'sec2_cargo_padrao':csc2, 'tes_padrao':ct}
                salvar_lote_configs(mudancas); st.rerun()

        with st.expander("🏢 Configs"):
            cn = st.text_input("Nome", db['config'].get('nome_conf',''))
            ch = st.text_input("Hora", db['config'].get('horario_padrao','20:00'))
            cl = st.text_input("Local", db['config'].get('local_padrao',''))
            cc = st.text_input("Cidade", db['config'].get('cidade_padrao',''))
            cpar = st.text_input("Cons. Part.", db['config'].get('cons_particular',''))
            ccen = st.text_input("Cons. Cent.", db['config'].get('cons_central',''))
            dfu = st.text_input("Dt Fund.", db['config'].get('data_fundacao',''))
            dag = st.text_input("Dt Agreg.", db['config'].get('data_agregacao',''))
            
            if st.button("Salvar Configs"):
                mudancas = {'nome_conf':cn, 'horario_padrao':ch, 'local_padrao':cl, 'cidade_padrao':cc, 
                            'cons_particular':cpar, 'cons_central':ccen, 'data_fundacao':dfu, 'data_agregacao':dag}
                salvar_lote_configs(mudancas); st.rerun()
                
        with st.expander("👥 Membros"):
            nm = st.text_input("Novo Membro")
            if st.button("Add"): gerenciar_lista_cloud("Membros","Nome",nm,"adicionar"); st.rerun()
            rm = st.selectbox("Remover", ["..."]+db['membros'])
            if st.button("Del") and rm != "...": gerenciar_lista_cloud("Membros","Nome",rm,"remover"); st.rerun()

        if st.button("Forçar Atualização"): limpar_memoria(); st.rerun()

    # === UI PRINCIPAL ===
    st.title("Gerador de Ata SSVP ✝️")
    
    val_num = int(dc.get('Numero', db['config']['ultima_ata'] + 1))
    val_data = data_pad
    if 'Data' in dc:
        try: val_data = datetime.strptime(dc['Data'], '%d/%m/%Y').date()
        except: pass

    c1, c2, c3 = st.columns(3)
    num_ata = c1.number_input("Número", value=val_num, step=1)
    if dc: st.caption(f"✏️ Editando Ata {val_num}")
    
    ia = db['anos'].index(dc['Ano']) if 'Ano' in dc and dc['Ano'] in db['anos'] else 0
    ano_tem = c2.selectbox("Ano Temático", db['anos'], index=ia)
    dt_reuniao = c3.date_input("Data", val_data, format="DD/MM/YYYY")

    with st.expander("📍 Detalhes da Reunião", expanded=False):
        cx1, cx2, cx3 = st.columns(3)
        hr_ini = cx1.time_input("Início", hora_pad)
        local_r = cx2.text_input("Local", db['config'].get('local_padrao',''))
        cidade_r = cx3.text_input("Cidade", db['config'].get('cidade_padrao',''))

    st.divider()
    st.subheader("Chamada")
    cp1, cp2 = st.columns(2)
    def_pres = [p.strip() for p in dc.get('Presentes','').split(',') if p.strip() in db['membros']]
    presentes = cp1.multiselect("1️⃣ Quem veio?", db['membros'], default=def_pres)
    ausentes = [m for m in db['membros'] if m not in presentes]
    motivos = {}
    justif = cp2.multiselect("2️⃣ Quem justificou?", ausentes)
    if justif:
        cols = st.columns(3)
        for i, m in enumerate(justif): motivos[m] = cols[i%3].text_input(m, placeholder="Motivo...")

    st.divider()
    st.subheader("Tesouraria")
    cf1, cf2, cf3, cf4 = st.columns(4)
    
    # Cache do saldo para não quebrar API
    if 'saldo_cache' not in st.session_state: st.session_state.saldo_cache = obter_saldo_anterior()
    saldo_ant = st.session_state.saldo_cache
    st.caption(f"Saldo Anterior: R$ {saldo_ant:.2f}")

    rec = cf1.number_input("Receita", value=float(dc.get('Receita', 0.0)), step=0.1)
    des = cf2.number_input("Despesa", value=float(dc.get('Despesa', 0.0)), step=0.1)
    dec = cf3.number_input("Décima", value=float(dc.get('Decima', 0.0)), step=0.1)
    saldo = cf4.number_input("Saldo Final", value=saldo_ant+rec-des-dec, disabled=True)
    tes_nome = cf4.selectbox("Tesoureiro", db['membros'], index=get_index_membro(db['config'].get('tes_padrao'), db['membros']))

    st.divider()
    ce1, ce2, ce3 = st.columns(3)
    pres_nome = ce1.selectbox("Presidente", db['membros'], index=get_index_membro(dc.get('Presidente', db['config'].get('pres_padrao')), db['membros']))
    font_l = ce2.text_input("Fonte Leitura", value=dc.get('Leitura',''))
    leit_nome = ce3.selectbox("Leitor", db['membros'])

    st.divider()
    st_ata = st.radio("Ata Anterior", ["Aprovada sem ressalvas", "Aprovada com ressalvas"])
    txt_res = st.text_input("Detalhes da ressalva") if st_ata == "Aprovada com ressalvas" else ""
    
    st.divider()
    visit = st.text_area("Visitantes", value=dc.get('Visitantes',''))
    st.markdown("### Relatórios")
    socio = st.text_area("Socioeconômico", value=dc.get('Socioeconomico',''))
    notic = st.text_area("Notícias", value=dc.get('Noticias',''))
    escal = st.text_area("Escala", value=dc.get('Escala',''))
    palav = st.text_area("Palavra Franca", value=dc.get('Palavra_Franca',''))
    exped = st.text_area("Expediente", value=dc.get('Expediente',''))
    
    st.divider()
    ce1, ce2 = st.columns(2)
    p_vis = ce1.text_input("Palavra Visitantes", "")
    mov_ex = ce2.text_input("Mov. Extra", "Coleta regular")
    ce3, ce4 = st.columns(2)
    music = ce3.text_input("Música", "Hino de Ozanam")
    hr_fim = ce4.time_input("Fim")

    st.divider()
    st.markdown("##### ✍️ Assinatura")
    qa = st.radio("Secretário Hoje?", ["1º Secretário", "2º Secretário", "Outro"], horizontal=True)
    if qa == "1º Secretário":
        idx_s = get_index_membro(db['config'].get('sec_padrao'), db['membros'])
        cg_fin = "1º Secretário(a)"
    elif qa == "2º Secretário":
        idx_s = get_index_membro(db['config'].get('sec2_padrao'), db['membros'])
        cg_fin = "2º Secretário(a)"
    else:
        idx_s = get_index_membro(dc.get('Secretario',''), db['membros'])
        cg_fin = "Secretário(a) ad hoc"
    sec_nom = st.selectbox("Nome Secretário", db['membros'], index=idx_s)

    st.divider()
    if st.button("💾 Gerar/Salvar Ata", type="primary"):
        ls_aus = [f"{m} ({motivos.get(m,'').strip() or 'Justificado'})" if m in motivos else m for m in ausentes]
        st_fin = f"{st_ata}: {txt_res}" if txt_res else st_ata
        
        dados_ata = {
            'num_ata': str(num_ata), 'conf_nome': db['config'].get('nome_conf',''),
            'cons_particular': db['config'].get('cons_particular',''), 'cons_central': db['config'].get('cons_central',''),
            'data_fundacao': formatar_data_br(db['config'].get('data_fundacao','')),
            'data_agregacao': formatar_data_br(db['config'].get('data_agregacao','')),
            'ano_tematico': ano_tem, 'data_reuniao': formatar_data_br(dt_reuniao),
            'hora_inicio': hr_ini.strftime('%H:%M'), 'local': local_r, 'pres_nome': pres_nome,
            'leitura_fonte': font_l, 'leitor_nome': leit_nome, 'status_ata_ant': st_fin,
            'lista_presentes_txt': ", ".join(presentes), 'ausencias': ", ".join(ls_aus) or "Não houve.",
            'lista_visitantes_txt': visit.replace("\n", ", ") if visit else "",
            'receita': rec, 'despesa': des, 'decima': dec, 'saldo': saldo,
            'tes_nome': tes_nome, 'socioeconomico': socio, 'noticias_trabalhos': notic,
            'escala_visitas': escal, 'palavra_franca': palav, 'expediente': exped,
            'palavra_visitantes': p_vis, 'mov_financeiro_extra': mov_ex,
            'musica_final': music, 'hora_fim': hr_fim.strftime('%H:%M'),
            'secretario_nome': sec_nom, 'secretario_cargo': cg_fin, 'cidade_estado': cidade_r
        }
        
        with st.spinner("Salvando..."):
            ok, tipo = salvar_historico_cloud(dados_ata)
            if ok:
                st.toast(f"✅ Ata {num_ata} salva!"); st.session_state.dados_carregados = {}
                if tipo=="criada" and int(num_ata) > db['config']['ultima_ata']:
                    salvar_lote_configs({'ultima_ata': int(num_ata)})
        
        doc = gerar_docx(dados_ata)
        bio = io.BytesIO(); doc.save(bio)
        pdf_bytes = gerar_pdf_nativo(dados_ata)
        
        c1, c2 = st.columns(2)
        c1.download_button("📄 PDF", pdf_bytes, f"Ata_{num_ata}.pdf", "application/pdf", type="primary")
        c2.download_button("📝 Word", bio.getvalue(), f"Ata_{num_ata}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

elif authentication_status == False: st.error("Login incorreto")
elif authentication_status == None: st.warning("Faça login")