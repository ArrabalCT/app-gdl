import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn
import io
import os
from PIL import Image, ImageOps

# --- Funções de Formatação do Word ---
def adicionar_borda_inferior(paragraph):
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'auto')
    pBdr.append(bottom)
    pPr.append(pBdr)

def adicionar_campo_numpages(paragraph):
    p = paragraph._p
    r1 = OxmlElement('w:r')
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    r1.append(fldChar1)
    p.append(r1)
    r2 = OxmlElement('w:r')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = ' NUMPAGES '
    r2.append(instrText)
    p.append(r2)
    r3 = OxmlElement('w:r')
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    r3.append(fldChar2)
    p.append(r3)
    r4 = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.text = 'xx'
    r4.append(t)
    p.append(r4)
    r5 = OxmlElement('w:r')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    r5.append(fldChar3)
    p.append(r5)

# --- Variáveis de Sessão ---
if 'fotos' not in st.session_state: st.session_state['fotos'] = []
if 'camera_key' not in st.session_state: st.session_state['camera_key'] = 0 
if 'mk' not in st.session_state: st.session_state['mk'] = 0 

mk = st.session_state['mk']

# --- Listas de Dados ---
delegados = ["Adilson Antonio Marcondes dos Santos", "Adriane Goncalves", "Anisio Galdioli", "Cesar Aparecido Vieira da Silva", "Daniel Souza Baptista de Castro", "Ernani Ronaldo Giannico Braga", "Fabio Germano Figueiredo Cabett", "Flavia Maria Rocha Rollo", "Francisco Sannini Neto", "Hugo Parreiras de Macedo", "Jose Marcelo Silva Hial", "Leonardo da Costa Ferreira", "Marcelo Vieira Cavalcante", "Mario Celso Ribeiro Senne", "Paulo Roberto Gruschka Castilho", "Paulo Sergio Barbosa", "Pedro Rossati", "Sergio Lucas Adler Guedes de Oliveira", "Vania Idalira Z. de Oliveira", "Outro..."]
peritos = ["Alexandre Rabello de Oliveira", "Bruna Fernandes Nogueira", "Claude Thiago Arrabal", "Jéssica Pereira Gonçalves", "Júlia Soares Melo", "Luiz Fausto Prado Vasques", "Marcelo Mourão Dantas", "Márcio Steinmetz Soares", "Sarah Costa Teixeira", "Ruan Carvalho de Souza"]
cidades = ["Aparecida", "Cachoeira Paulista", "Canas", "Cunha", "Guaratinguetá", "Lorena", "Piquete", "Potim", "Roseira", "Outra..."]
dps_por_cidade = {"Aparecida": ["DEL.POL.APARECIDA"], "Canas": ["DEL.POL.CANAS"], "Cachoeira Paulista": ["DEL.POL.CACHOEIRA PAULISTA"], "Cunha": ["DEL.POL.CUNHA"], "Guaratinguetá": ["01º D.P. GUARATINGUETA", "02º D.P. GUARATINGUETA", "DEL.SEC.GUARATINGUETA PLANTÃO", "DISE- DEL.SEC.GUARATINGUETA"], "Lorena": ["01º D.P. LORENA", "02º D.P. LORENA", "DEL.POL.LORENA"], "Piquete": ["DEL.POL.PIQUETE"], "Potim": ["DEL.POL.POTIM"], "Roseira": ["DEL.POL.ROSEIRA"]}

# --- INTERFACE ---
st.title("Gerador de Laudos - Celular")

st.header("1. Cabeçalho e Identificação")

colR1, colR2 = st.columns(2)
with colR1: 
    rep_input = st.text_input("Número REP:", value="", placeholder="Ex: 12345", key=f"rep_{mk}")
with colR2: 
    rep_ano = st.text_input("Ano REP:", value="2026", max_chars=4, key=f"rep_ano_{mk}")

colBO1, colBO2 = st.columns(2)
with colBO1: 
    bo_input = st.text_input("Número do BO:", value="", placeholder="Ex: LT0644", key=f"bo_{mk}").upper()
with colBO2: 
    bo_ano = st.text_input("Ano do BO:", value="2026", max_chars=4, key=f"ano_{mk}")

data_selecionada = st.date_input("Data do Laudo:", format="DD/MM/YYYY", key=f"data_{mk}")
meses = {1: 'janeiro', 2: 'fevereiro', 3: 'março', 4: 'abril', 5: 'maio', 6: 'junho', 7: 'julho', 8: 'agosto', 9: 'setembro', 10: 'outubro', 11: 'novembro', 12: 'dezembro'}
data_extenso = f"{data_selecionada.day} de {meses[data_selecionada.month]} de {data_selecionada.year}"

objetivos_selecionados = st.multiselect("Objetivo da Perícia:", ["Extração de Dados - Cellebrite", "Constatação de danos", "Fotografação"], default=["Extração de Dados - Cellebrite", "Constatação de danos"], key=f"obj_{mk}")
perito_selecionado = st.selectbox("Perito Criminal:", peritos, index=peritos.index("Claude Thiago Arrabal"), key=f"per_{mk}")

del_sel = st.selectbox("Autoridade Policial:", delegados, index=delegados.index("Adilson Antonio Marcondes dos Santos") if "Adilson Antonio Marcondes dos Santos" in delegados else 0, key=f"del_sel_{mk}")
if del_sel == "Outro...":
    delegado_selecionado = st.text_input("Digite o nome da Autoridade Policial:", key=f"del_dig_{mk}")
else:
    delegado_selecionado = del_sel

colC1, colC2 = st.columns(2)
with colC1: 
    cid_sel = st.selectbox("Cidade:", cidades, index=cidades.index("Guaratinguetá") if "Guaratinguetá" in cidades else 0, key=f"cid_sel_{mk}")
    if cid_sel == "Outra...":
        cidade_selecionada = st.text_input("Digite o nome da Cidade:", key=f"cid_dig_{mk}")
    else:
        cidade_selecionada = cid_sel

with colC2:
    if cid_sel == "Outra...":
        delegacia_selecionada = st.text_input("Digite o nome da Delegacia:", key=f"dp_dig_{mk}")
    else:
        opcoes_dp = dps_por_cidade[cid_sel] + ["Outra..."]
        index_padrao = opcoes_dp.index("DEL.SEC.GUARATINGUETA PLANTÃO") if "DEL.SEC.GUARATINGUETA PLANTÃO" in opcoes_dp else 0
        dp_sel = st.selectbox("Delegacia:", opcoes_dp, index=index_padrao, key=f"dp_sel_{mk}")
        if dp_sel == "Outra...":
            delegacia_selecionada = st.text_input("Digite o nome da Delegacia:", key=f"dp_dig_esp_{mk}")
        else:
            delegacia_selecionada = dp_sel

st.header("2. Cadeia de Custódia e Aparelhos")

lacre_saida = st.text_input("Lacre de Saída (Envio para o NI):", placeholder="Ex: 0012345", key=f"lacre_saida_{mk}")
qtd_aparelhos = st.number_input("Quantidade de Aparelhos Recebidos:", min_value=1, max_value=20, value=1, step=1, key=f"qtd_{mk}")

dados_aparelhos = []

# --- LOOP PARA MÚLTIPLOS CELULARES ---
for i in range(qtd_aparelhos):
    st.markdown(f"### 📱 Item {i+1}")
    lacre_ent = st.text_input(f"Lacre de Entrada do Item {i+1}:", placeholder="Ex: 9876543", key=f"lacre_ent_{i}_{mk}")
    
    col1, col2 = st.columns(2)
    with col1:
        tipo_opcoes = sorted(["Smartphone", "Tablet", "Feature Phone", "Smartwatch"]) + ["Outro"]
        tipo_sel = st.selectbox(f"Tipo (Item {i+1}):", tipo_opcoes, key=f"tipo_sel_{i}_{mk}")
        tipo_cel = st.text_input(f"Qual tipo? (Item {i+1})", key=f"tipo_dig_{i}_{mk}") if tipo_sel == "Outro" else tipo_sel

        marca_opcoes = sorted(["Motorola", "Samsung", "Apple", "Xiaomi", "LG", "Asus", "Nokia", "Positivo", "Realme", "Multilaser"]) + ["Outra"]
        marca_sel = st.selectbox(f"Marca (Item {i+1}):", marca_opcoes, key=f"marca_sel_{i}_{mk}")
        marca_cel = st.text_input(f"Qual marca? (Item {i+1})", key=f"marca_dig_{i}_{mk}") if marca_sel == "Outra" else marca_sel

        modelo_cel = st.text_input(f"Modelo (Item {i+1}):", placeholder="Se vazio = não aparente", key=f"mod_{i}_{mk}")

    with col2:
        cor_opcoes = sorted(["Preta", "Branca", "Prata", "Cinza", "Azul", "Dourada", "Rosa", "Vermelha", "Amarela", "Verde", "Roxa"]) + ["Outra"]
        cor_sel = st.selectbox(f"Cor (Item {i+1}):", cor_opcoes, key=f"cor_sel_{i}_{mk}")
        cor_cel = st.text_input(f"Qual cor? (Item {i+1})", key=f"cor_dig_{i}_{mk}") if cor_sel == "Outra" else cor_sel

        imei_cel = st.text_input(f"IMEI (Item {i+1}):", placeholder="Se vazio = não aparente", key=f"imei_{i}_{mk}")
        tem_capa = st.radio(f"Capa? (Item {i+1})", ["Sim", "Não"], horizontal=True, key=f"capa_{i}_{mk}")

    st.markdown(f"**SIMCards (Chips) - Item {i+1}**")
    sim_opcoes = sorted(["Vivo", "Tim", "Claro", "Oi", "Correios Cellular"]) + ["Outra", "Nenhum"]

    c_sim1, c_sim2 = st.columns(2)
    with c_sim1:
        s1_sel = st.selectbox(f"Operadora 1 (Item {i+1}):", sim_opcoes, index=sim_opcoes.index("Nenhum"), key=f"s1_sel_{i}_{mk}")
        s1_txt = st.text_input(f"Qual operadora 1? (Item {i+1})", key=f"s1_dig_{i}_{mk}") if s1_sel == "Outra" else s1_sel
        iccid1 = st.text_input(f"ICCID 1 (Item {i+1}):", key=f"icc1_{i}_{mk}")

    with c_sim2:
        s2_sel = st.selectbox(f"Operadora 2 (Item {i+1}):", sim_opcoes, index=sim_opcoes.index("Nenhum"), key=f"s2_sel_{i}_{mk}")
        s2_txt = st.text_input(f"Qual operadora 2? (Item {i+1})", key=f"s2_dig_{i}_{mk}") if s2_sel == "Outra" else s2_sel
        iccid2 = st.text_input(f"ICCID 2 (Item {i+1}):", key=f"icc2_{i}_{mk}")

    st.markdown(f"**Danos - Item {i+1}**")
    locais_opcoes = ["Tela (Display)", "Tampa traseira", "Lentes da câmera", "Botões", "Bordas/Laterais", "Película de proteção"]
    locais_selecionados = st.multiselect(f"Onde há danos? (Item {i+1}):", locais_opcoes, key=f"loc_{i}_{mk}")

    danos_detalhes = {}
    extensoes_opcoes = ["Toda a extensão", "Canto superior esquerdo", "Canto superior direito", "Canto inferior esquerdo", "Canto inferior direito", "Centro", "Margens", "Terço superior", "Terço inferior", "Outro..."]
    
    for local in locais_selecionados:
        cd1, cd2 = st.columns(2)
        with cd1:
            t_dano = st.multiselect(f"Dano ({local}) - Item {i+1}:", ["Fratura", "Quebra", "Trinco", "Riscos", "Atritamento", "Mancha"], key=f"td_{local}_{i}_{mk}")
        with cd2:
            ext_sel = st.multiselect(f"Extensão ({local}) - Item {i+1}:", extensoes_opcoes, key=f"ext_sel_{local}_{i}_{mk}")
            ext_txt = ""
            if "Outro..." in ext_sel:
                ext_txt = st.text_input(f"Digitar Extensão ({local}):", key=f"ext_txt_{local}_{i}_{mk}")
        danos_detalhes[local] = {"tipo": t_dano, "ext_sel": ext_sel, "ext_txt": ext_txt}

    # Salvando os dados deste item no dicionário geral
    dados_aparelhos.append({
        "lacre": lacre_ent, "tipo": tipo_cel, "marca": marca_cel, "modelo": modelo_cel, 
        "cor": cor_cel, "imei": imei_cel, "capa": tem_capa, "s1": s1_txt, "icc1": iccid1,
        "s2": s2_txt, "icc2": iccid2, "danos": locais_selecionados, "detalhes": danos_detalhes
    })
    st.markdown("---")


# --- GERADOR DE TEXTO ---
txt_gerado = "Foi recebido para exame, lacrado(s) em sacos plásticos transparentes, os seguintes itens:\n\n"

for idx, ap in enumerate(dados_aparelhos):
    lacre_str = ap['lacre'] if ap['lacre'].strip() else "não informado"
    txt_gerado += f"Item {idx+1} --- Lacre de entrada {lacre_str}\n"
    
    mod_str = ap['modelo'].upper() if ap['modelo'] else "não aparente"
    capa_str = "acompanha capa de proteção" if ap['capa'] == "Sim" else "não acompanha capa de proteção"
    imei_str = f"IMEI {ap['imei']}" if ap['imei'].strip() else "IMEI não aparente"

    # Lógica dos SIMCards
    sims_encontrados = []
    if ap['s1'] != "Nenhum":
        txt_1 = f"da operadora {ap['s1']}"
        if ap['icc1'].strip(): txt_1 += f" (ICCID {ap['icc1']})"
        sims_encontrados.append(txt_1)

    if ap['s2'] != "Nenhum":
        txt_2 = f"da operadora {ap['s2']}"
        if ap['icc2'].strip(): txt_2 += f" (ICCID {ap['icc2']})"
        sims_encontrados.append(txt_2)

    if not sims_encontrados:
        sim_str = "não acompanha SIMCard"
    else:
        sim_str = "acompanhado de SIMCard(s) " + " e ".join(sims_encontrados)

    txt_gerado += f"Trata-se de um aparelho do tipo {ap['tipo'].lower()}, marca {ap['marca'].upper()}, modelo {mod_str}, de cor predominante {ap['cor'].lower()}, que {capa_str}. O dispositivo apresenta {imei_str} e {sim_str}.\n"

    # Lógica dos Danos
    if not ap['danos']:
        txt_gerado += "Ao exame físico externo, o dispositivo não apresentava danos ou avarias visíveis em seu display ou carcaça.\n\n"
    else:
        linhas_dano = []
        for loc in ap['danos']:
            det = ap['detalhes'][loc]
            tipos_str = ", ".join(det["tipo"]).lower() if det["tipo"] else "avarias"
            
            # Montando string de extensão
            lista_ext = [e for e in det["ext_sel"] if e != "Outro..."]
            if det["ext_txt"].strip(): lista_ext.append(det["ext_txt"].strip())
            ext_str = f" ({', '.join(lista_ext).lower()})" if lista_ext else ""
            
            linhas_dano.append(f"{loc.lower()} com {tipos_str}{ext_str}")
        
        txt_gerado += f"Ao exame físico externo, constatou-se a presença das seguintes constatações: {'; '.join(linhas_dano)}.\n\n"

# Fechamento com Lacre de Saída
lsaida_str = lacre_saida if lacre_saida.strip() else "[INSERIR LACRE DE SAÍDA]"
txt_gerado += f"O(s) aparelho(s) foram enviados para extração de dados no Núcleo de Informática de São José dos Campos em lacre {lsaida_str}. O resultado seguirá em laudo complementar."

if st.session_state.get(f"track_{mk}") != txt_gerado:
    st.session_state[f"edit_{mk}"] = txt_gerado
    st.session_state[f"track_{mk}"] = txt_gerado

st.header("3. Edição e Word")
st.warning("⚠️ Pode digitar diretamente na caixa abaixo. No entanto, deixe as edições manuais para o **FINAL**. Se alterar as opções em cima, o sistema irá recriar a frase e apagar as suas edições!")
texto_final = st.text_area("Texto final que vai para o Laudo:", height=400, key=f"edit_{mk}")

# FOTOS 
foto = st.camera_input("Tirar fotografia", key=f"cam_{st.session_state['camera_key']}")
if foto:
    colA, colR = st.columns(2)
    with colA:
        if st.button("✅ ACEITAR FOTO", type="primary", use_container_width=True):
            img = Image.open(io.BytesIO(foto.getvalue()))
            st.session_state['fotos'].append({'img': ImageOps.exif_transpose(img), 'nome': f"cam_{st.session_state['camera_key']}"})
            st.session_state['camera_key'] += 1
            st.rerun()
    with colR:
        if st.button("❌ REJEITAR", use_container_width=True):
            st.session_state['camera_key'] += 1
            st.rerun()

fotos_up = st.file_uploader("Ou carregue da galeria", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True, key=f"up_{mk}")
if fotos_up:
    nomes_existentes = [f.get('nome') for f in st.session_state['fotos'] if 'nome' in f]
    for f in fotos_up:
        if f.name not in nomes_existentes:
            img = Image.open(io.BytesIO(f.getvalue()))
            st.session_state['fotos'].append({'img': ImageOps.exif_transpose(img), 'nome': f.name})

if st.session_state['fotos']:
    st.markdown("### 📸 Fotos Anexadas")
    cols = st.columns(3)
    for i, foto_data in enumerate(st.session_state['fotos']):
        with cols[i % 3]:
            st.image(foto_data['img'], use_container_width=True)
            if st.button("❌ Apagar", key=f"del_{i}_{mk}", type="secondary"):
                st.session_state['fotos'].pop(i)
                st.rerun()

# --- FINALIZAÇÃO E DOWNLOAD ---
st.header("4. Finalizar")
c1, c2 = st.columns(2)

with c1:
    if st.button("Criar Laudo (.docx)", type="primary", use_container_width=True):
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Courier New'
        style.font.size = Pt(11)
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        # Cabeçalho
        section = doc.sections[0]
        header = section.header
        for p in header.paragraphs: p.text = ""
        
        table = header.add_table(rows=1, cols=3, width=Cm(15.5))
        table.autofit = False
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        largura_lateral = Cm(2.2)
        largura_meio = Cm(11.1)

        table.columns[0].width = largura_lateral
        table.columns[1].width = largura_meio
        table.columns[2].width = largura_lateral
        
        for cell in table.columns[0].cells: cell.width = largura_lateral
        for cell in table.columns[1].cells: cell.width = largura_meio
        for cell in table.columns[2].cells: cell.width = largura_lateral

        p_left = table.cell(0, 0).paragraphs[0]; p_left.alignment = WD_ALIGN_PARAGRAPH.LEFT
        if os.path.exists("logo_ssp.png"): p_left.add_run().add_picture("logo_ssp.png", width=Cm(1.8))
        
        p_center = table.cell(0, 1).paragraphs[0]; p_center.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_h1 = p_center.add_run("SECRETARIA DA SEGURANÇA PÚBLICA\nSUPERINTENDÊNCIA DA POLÍCIA TÉCNICO-CIENTÍFICA\n")
        run_h1.font.size = Pt(11); run_h1.bold = False
        run_h2 = p_center.add_run("INSTITUTO DE CRIMINALÍSTICA\n“PERITO CRIMINAL DR. OCTÁVIO EDUARDO DE BRITO ALVARENGA”\nNÚCLEO DE PERÍCIAS CRIMINALÍSTICAS DE SÃO JOSÉ DOS CAMPOS\nEQUIPE DE PERÍCIAS CRIMINALÍSTICAS DE GUARATINGUETÁ")
        run_h2.font.size = Pt(8); run_h2.bold = False

        p_right = table.cell(0, 2).paragraphs[0]; p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if os.path.exists("logo_ic.png"): p_right.add_run().add_picture("logo_ic.png", width=Cm(1.8))

        if rep_input or bo_input:
            p_ident = doc.add_paragraph()
            p_ident.alignment = WD_ALIGN_PARAGRAPH.CENTER
            ident_text = []
            if bo_input: ident_text.append(f"BO {bo_input} / {bo_ano} - {delegacia_selecionada}")
            p_ident.add_run(" | ".join(ident_text))

        # Corpo
        p_nat = doc.add_paragraph()
        run = p_nat.add_run("1 – NATUREZA: Peças"); run.bold = True; run.font.size = Pt(14)
        adicionar_borda_inferior(p_nat)
        
        preambulo = (f"Aos {data_extenso}, no Instituto de Criminalística, da Superintendência da Polícia Técnico-Científica, "
                     f"da Secretaria da Segurança Pública do Estado de São Paulo, de conformidade com o disposto no artigo 178 "
                     f"do Decreto-Lei nº. 3689, de 03 de outubro de 1941, pelo Diretor do Instituto de Criminalística, Ricardo Lopes Ortega, "
                     f"foi designado o Perito Criminal {perito_selecionado}, para proceder ao exame supracitado, em atendimento à requisição "
                     f"da Autoridade Policial, Dr(a). {delegado_selecionado}, titular/em exercício na {delegacia_selecionada}.")
        doc.add_paragraph(preambulo)

        p_obj = doc.add_paragraph()
        run_obj = p_obj.add_run("2 - OBJETIVO DA PERÍCIA:"); run_obj.bold = True; run_obj.font.size = Pt(14)
        adicionar_borda_inferior(p_obj)
        objetivos_str = ", ".join(objetivos_selecionados) if objetivos_selecionados else "Não especificado"
        doc.add_paragraph(f"Consta na requisição de exame: {objetivos_str}.")

        p_ex = doc.add_paragraph()
        run2 = p_ex.add_run("3 – DOS EXAMES:"); run2.bold = True; run2.font.size = Pt(14)
        adicionar_borda_inferior(p_ex)
        
        # Inserindo o texto formatado mantendo parágrafos vazios
        for linha in texto_final.split("\n"): 
            if linha.strip() == "":
                doc.add_paragraph("")
            else:
                doc.add_paragraph(linha.strip())
            
        if st.session_state['fotos']:
            doc.add_page_break()
            p_fotos = doc.add_paragraph()
            run_fotos = p_fotos.add_run("4 - REGISTRO FOTOGRÁFICO"); run_fotos.bold = True; run_fotos.font.size = Pt(14)
            adicionar_borda_inferior(p_fotos)
            
            for i, foto_data in enumerate(st.session_state['fotos']):
                img = foto_data['img']
                if img.mode != 'RGB': img = img.convert('RGB')
                largura_foto = Cm(14.0) if img.width > img.height else Cm(9.5)
                buf = io.BytesIO()
                img.save(buf, format='JPEG', quality=90)
                buf.seek(0)
                
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.add_run().add_picture(buf, width=largura_foto)
                
                legenda = doc.add_paragraph(f"Fotografia {i+1}: Dispositivo examinado.")
                legenda.alignment = WD_ALIGN_PARAGRAPH.CENTER
                doc.add_paragraph("")
            
        # Encerramento
        p_relatar = doc.add_paragraph("Era o que havia a relatar.")
        p_relatar.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        p_paginas = doc.add_paragraph("Este laudo vai impresso em ")
        adicionar_campo_numpages(p_paginas)  
        p_paginas.add_run(" páginas, além da capa, ficando arquivada cópia digital no sistema GDL da SPTC.")
        p_paginas.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        p_assinatura = doc.add_paragraph(); p_assinatura.paragraph_format.space_after = p_assinatura.paragraph_format.space_before = Pt(0)
        p_assinatura.add_run(perito_selecionado).bold = True
        p_assinatura.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        p_cargo = doc.add_paragraph("Perito Criminal Relator"); p_cargo.paragraph_format.space_after = p_cargo.paragraph_format.space_before = Pt(0)
        p_cargo.alignment = WD_ALIGN_PARAGRAPH.CENTER

        buf_doc = io.BytesIO(); doc.save(buf_doc); buf_doc.seek(0)
        
        if rep_input:
            nome_arquivo = f"Laudo_Cel_REP_{rep_input}_{rep_ano}.docx"
        elif bo_input:
            nome_arquivo = f"Laudo_Cel_BO_{bo_input}_{bo_ano}.docx"
        else:
            nome_arquivo = "Laudo_Cel_Sem_BO_REP.docx"
            
        st.download_button("⬇️ Descarregar Laudo Final", buf_doc, nome_arquivo, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

with c2:
    if st.button("🔄 Novo(s) Celular(es) (Limpar Tudo)", type="secondary", use_container_width=True):
        current_mk = st.session_state.get('mk', 0)
        st.session_state.clear()
        st.session_state['mk'] = current_mk + 1
        st.rerun()