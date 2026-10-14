from pathlib import Path
import re, shutil, unicodedata
from PIL import Image
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).parent
SRC = ROOT / 'Manual_IPT_4a_Edicao.docx'
OUT = ROOT / 'Manual_IPT_4a_Edicao_com_imagens.docx'
ASSET_DIR = ROOT / 'build_assets'
ASSET_DIR.mkdir(exist_ok=True)

# Export compact, high-resolution copies suitable for a PDF while preserving legibility.
def optimized_image(src: Path) -> Path:
    out = ASSET_DIR / (src.stem + '.jpg')
    if not out.exists() or out.stat().st_mtime < src.stat().st_mtime:
        im = Image.open(src).convert('RGB')
        im.thumbnail((1500, 1150), Image.Resampling.LANCZOS)
        im.save(out, 'JPEG', quality=88, optimize=True, progressive=True)
    return out

def norm(s):
    s = unicodedata.normalize('NFKD', s).encode('ascii','ignore').decode().lower()
    return re.sub(r'[^a-z0-9]+', ' ', s).strip()

def parse_mapping():
    rows=[]
    text=(ROOT/'manual_images'/'MAPEAMENTO_IMAGENS.md').read_text()
    for line in text.splitlines():
        m=re.match(r'\|\s*`([^`]+)`\s*\|\s*([^|]+)\|\s*([^|]+)\|', line)
        if m:
            rows.append((ROOT/'manual_images'/m.group(1), m.group(2).strip(), m.group(3).strip()))
    return rows

def insert_after(paragraph, image, caption):
    p = paragraph._p
    img_p = OxmlElement('w:p')
    p.addnext(img_p)
    from docx.text.paragraph import Paragraph
    ip = Paragraph(img_p, paragraph._parent)
    ip.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = ip.add_run()
    r.add_picture(str(image), width=Inches(6.15))
    cp = OxmlElement('w:p')
    img_p.addnext(cp)
    cap = Paragraph(cp, paragraph._parent)
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rr=cap.add_run(caption)
    rr.italic=True; rr.font.size=Pt(8); rr.font.name='Arial'
    return cap

def add_heading_before_body(doc, text):
    body=doc._body._element
    sect=body.sectPr
    p=OxmlElement('w:p')
    if sect is not None: body.insert(body.index(sect), p)
    else: body.append(p)
    from docx.text.paragraph import Paragraph
    para=Paragraph(p, doc._body)
    para.style='Heading 1'
    para.add_run(text)
    return para

def heading_candidates(doc):
    out=[]
    for p in doc.paragraphs:
        txt=p.text.strip()
        if not txt: continue
        st=p.style.name if p.style else ''
        if st.startswith('Heading') or re.match(r'^(Unidade|Unidades)\b',txt):
            out.append((p,txt,st))
    return out

def choose_anchor(location, candidates, path):
    loc=location.replace('—','-').replace('–','-')
    # Prefer the numbered section before the em dash; this is stable in the DOCX.
    m=re.search(r'\b(\d+(?:\.\d+)*)', loc)
    sec=m.group(1) if m else None
    ln=norm(loc)
    keywords=[x for x in ln.split() if len(x)>3 and x not in {'unidade','unidades','continuação','cont'}]
    scored=[]
    for p,txt,st in candidates:
        nt=norm(txt)
        score=0
        # Section numbers are rendered with dots in headings but become spaces
        # under accent/punctuation normalization, so test the original text too.
        if sec and (re.match(r'^\s*'+re.escape(sec)+r'(?:[.\s])', txt) or re.search(r'\b'+re.escape(sec.replace('.', r'\D+'))+r'\b', txt)):
            score += 30
        score += sum(1 for k in keywords if k in nt)
        if ' / ' in location and sec and re.match(r'^\s*'+re.escape(sec)+r'(?:[.\s])', txt): score += 8
        scored.append((score,p,txt))
    scored.sort(key=lambda x:x[0], reverse=True)
    if scored and scored[0][0] >= (30 if sec else 2): return scored[0][1]
    # Intro cover goes after Nota Introdutória.
    if 'Capa' in location:
        for p,t,s in candidates:
            if norm(t)=='nota introdutoria': return p
    return None

# Explicit anchors for the supplementary flowcharts/illustrations.
FLOW_ANCHORS={
 'fluxo_artigos_indefinidos.png':'2.1.2.', 'fluxo_condicionais.png':'16.3.',
 'fluxo_conectores_escolha.png':'32.4.', 'fluxo_demonstrativos.png':'2.1.1.',
 'fluxo_discurso_indireto.png':'17.1.', 'fluxo_enfase_cleft.png':'25.2.',
 'fluxo_escrita_academica.png':'33', 'fluxo_gerundio_infinitivo.png':'16.2.',
 'fluxo_inversao.png':'25.3.', 'fluxo_modais_escolha.png':'10.1.',
 'fluxo_noun_clauses.png':'32.3.', 'fluxo_participios_ed_ing.png':'27.3.',
 'fluxo_plural_substantivos.png':'2.1.4.', 'fluxo_preposicoes_tempo.png':'11.2.',
 'fluxo_pronomes_relativos.png':'14.2.', 'fluxo_tempos_passado.png':'9.2.',
 'fluxo_tempos_presente.png':'4.1.', 'fluxo_voz_passiva.png':'13.1.',
 'ilustra_collocations_do_make.png':'33', 'ilustra_estrutura_ensaio.png':'33',
 'ilustra_paralelismo.png':'28.4.', 'ilustra_phrasal_business.png':'14.3.',
 'ilustra_pontuacao_academica.png':'33', 'ilustra_prefixos_negativos.png':'26.4.'
}

def find_by_prefix(prefix,candidates):
    for p,t,s in candidates:
        if norm(t).startswith(norm(prefix)) or re.search(r'\b'+re.escape(prefix).replace('\\.','\\.')+r'\b', norm(t)):
            return p
    # numeric token match
    n=norm(prefix).rstrip('.')
    for p,t,s in candidates:
        if n and re.search(r'\b'+re.escape(n)+r'\b',norm(t)): return p
    return None

def main():
    doc=Document(str(SRC))
    candidates=heading_candidates(doc)
    placements=[]; used=set()
    rows=parse_mapping()
    for src,topic,location in rows:
        if not src.exists(): continue
        anchor=choose_anchor(location,candidates,src)
        if anchor is not None:
            placements.append((anchor,src,topic)); used.add(src.resolve())
    # Supplementary flowcharts and illustrations, mapped to exact section prefixes.
    for src in sorted((ROOT/'manual_fluxogramas').glob('*.png')):
        anchor=find_by_prefix(FLOW_ANCHORS.get(src.name,''), candidates)
        if anchor is not None:
            placements.append((anchor,src,src.stem.replace('_',' ').title())); used.add(src.resolve())
    # Insert in reverse document order so XML anchors remain stable.
    order={id(p):i for i,p in enumerate(doc.paragraphs)}
    placements.sort(key=lambda x: order.get(id(x[0]),0), reverse=True)
    for anchor,src,topic in placements:
        img=optimized_image(src)
        insert_after(anchor,img, f'Figura — {topic}')
    # Any image not anchored gets a complete gallery at the end; this guarantees 100% coverage.
    all_imgs=sorted([p for d in ('manual_images','manual_fluxogramas') for p in (ROOT/d).rglob('*.png')])
    missing=[p for p in all_imgs if p.resolve() not in used]
    if missing:
        h=add_heading_before_body(doc,'Galeria complementar de imagens')
        q=h
        for src in missing:
            q=insert_after(q,optimized_image(src),f'Imagem complementar — {src.stem.replace("_"," ").title()}')
    # Update core properties.
    doc.core_properties.title='Manual IPT — 4ª Edição (Livro Ilustrado)'
    doc.core_properties.subject='Manual de Inglês para Propósitos Técnicos com imagens didáticas'
    doc.core_properties.comments=f'Composição automática: {len(all_imgs)} imagens incorporadas.'
    doc.save(str(OUT))
    print(f'rows={len(rows)} anchored={len(placements)} total={len(all_imgs)} missing_gallery={len(missing)}')
    print(f'output={OUT}')

if __name__=='__main__': main()
