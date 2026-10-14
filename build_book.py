from pathlib import Path
import re
from collections import defaultdict
from PIL import Image
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

ROOT = Path(__file__).parent
SRC = ROOT / 'Manual_IPT_4a_Edicao.docx'
OUT = ROOT / 'Manual_IPT_4a_Edicao_com_imagens.docx'
ASSET_DIR = ROOT / 'build_assets'
ASSET_DIR.mkdir(exist_ok=True)


def optimized_image(src: Path) -> Path:
    """Create a compact copy for the document without sacrificing page readability."""
    out = ASSET_DIR / (src.stem + '.jpg')
    if not out.exists() or out.stat().st_mtime < src.stat().st_mtime:
        im = Image.open(src).convert('RGB')
        im.thumbnail((1500, 1150), Image.Resampling.LANCZOS)
        im.save(out, 'JPEG', quality=88, optimize=True, progressive=True)
    return out


def parse_mapping():
    rows = []
    text = (ROOT / 'manual_images' / 'MAPEAMENTO_IMAGENS.md').read_text()
    for line in text.splitlines():
        m = re.match(r'\|\s*`([^`]+)`\s*\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|', line)
        if m:
            rows.append((ROOT / 'manual_images' / m.group(1), m.group(2).strip(), m.group(3).strip()))
    return rows


# Supplementary flowcharts/illustrations are mapped to the exact numbered topic.
FLOW_ANCHORS = {
    'fluxo_artigos_indefinidos.png': '2.1.2.',
    'fluxo_condicionais.png': '16.3.',
    'fluxo_conectores_escolha.png': '32.4.',
    'fluxo_demonstrativos.png': '2.1.1.',
    'fluxo_discurso_indireto.png': '17.1.',
    'fluxo_enfase_cleft.png': '25.2.',
    'fluxo_escrita_academica.png': '33',
    'fluxo_gerundio_infinitivo.png': '16.2.',
    'fluxo_inversao.png': '25.3.',
    'fluxo_modais_escolha.png': '10.1.',
    'fluxo_noun_clauses.png': '32.3.',
    'fluxo_participios_ed_ing.png': '27.3.',
    'fluxo_plural_substantivos.png': '2.1.4.',
    'fluxo_preposicoes_tempo.png': '11.2.',
    'fluxo_pronomes_relativos.png': '14.2.',
    'fluxo_tempos_passado.png': '9.2.',
    'fluxo_tempos_presente.png': '4.1.',
    'fluxo_voz_passiva.png': '13.1.',
    'ilustra_collocations_do_make.png': '33',
    'ilustra_estrutura_ensaio.png': '33',
    'ilustra_paralelismo.png': '28.4.',
    'ilustra_phrasal_business.png': '14.3.',
    'ilustra_pontuacao_academica.png': '33',
    'ilustra_prefixos_negativos.png': '26.4.',
}


def heading_level(p):
    # Word's TOC styles are named e.g. "toc 3"; they must not be treated as
    # content headings, otherwise figures would also be inserted in the index.
    m = re.match(r'^Heading\s*(\d+)$', p.style.name if p.style else '')
    return int(m.group(1)) if m else None


def real_headings(doc):
    # Only Word heading styles are used. This deliberately excludes the TOC,
    # whose entries look like headings but are not content sections.
    return [(i, p, heading_level(p)) for i, p in enumerate(doc.paragraphs)
            if heading_level(p) is not None and p.text.strip()]


def section_number(location):
    m = re.search(r'\b(\d+(?:\.\d+)*)\b', location)
    return m.group(1) if m else None


def number_matches(text, sec):
    if not sec:
        return False
    # Match a true heading prefix, e.g. 1.2.2. Do not match 1.2.20.
    return bool(re.match(r'^\s*' + re.escape(sec) + r'(?:\.|\s|$)', text))


def find_topic_heading(location, headings):
    sec = section_number(location)
    if sec:
        exact = [(i, p, level) for i, p, level in headings if number_matches(p.text, sec)]
        if exact:
            return exact[0]
        # Unit 33 and similar destinations have a number embedded in the unit title.
        embedded = [(i, p, level) for i, p, level in headings
                    if re.search(r'\b' + re.escape(sec) + r'\b', p.text)]
        if embedded:
            return embedded[0]
    if 'Capa' in location:
        for item in headings:
            if item[1].text.strip().lower() == 'nota introdutória':
                return item
    return None


def find_numbered_heading(prefix, headings):
    if prefix == '33':
        for item in headings:
            if re.search(r'\b33\b', item[1].text):
                return item
    for item in headings:
        if number_matches(item[1].text, prefix):
            return item
    return None


def next_boundary(target_index, target_level, headings):
    """Return the next same-or-higher-level heading after a topic."""
    for i, p, level in headings:
        if i > target_index and level <= target_level:
            return p
    return None


def add_figure_before(boundary, parent, image, caption):
    """Insert an image and caption immediately before the boundary paragraph."""
    image_p = OxmlElement('w:p')
    caption_p = OxmlElement('w:p')
    if boundary is None:
        body = parent._element.body
        sect = body.sectPr
        if sect is not None:
            body.insert(body.index(sect), image_p)
            body.insert(body.index(sect), caption_p)
        else:
            body.append(image_p)
            body.append(caption_p)
    else:
        boundary._p.addprevious(image_p)
        boundary._p.addprevious(caption_p)
    ip = Paragraph(image_p, parent)
    ip.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ip.paragraph_format.keep_with_next = True
    ip.add_run().add_picture(str(image), width=Inches(6.15))
    cp = Paragraph(caption_p, parent)
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.paragraph_format.keep_together = True
    run = cp.add_run(caption)
    run.italic = True
    run.font.size = Pt(8)
    run.font.name = 'Arial'


def main():
    doc = Document(str(SRC))
    headings = real_headings(doc)
    placements = []
    used = set()
    unmatched = []

    for src, topic, location in parse_mapping():
        if not src.exists():
            continue
        target = find_topic_heading(location, headings)
        if target is None:
            unmatched.append((src.name, location))
            continue
        target_i, target_p, target_level = target
        boundary = next_boundary(target_i, target_level, headings)
        placements.append((boundary, src, topic, target_p.text.strip()))
        used.add(src.resolve())

    for src in sorted((ROOT / 'manual_fluxogramas').glob('*.png')):
        prefix = FLOW_ANCHORS.get(src.name)
        target = find_numbered_heading(prefix, headings) if prefix else None
        if target is None:
            unmatched.append((src.name, prefix or 'sem destino'))
            continue
        target_i, target_p, target_level = target
        boundary = next_boundary(target_i, target_level, headings)
        placements.append((boundary, src, src.stem.replace('_', ' ').title(), target_p.text.strip()))
        used.add(src.resolve())

    all_imgs = sorted([p for d in ('manual_images', 'manual_fluxogramas')
                       for p in (ROOT / d).rglob('*.png')])
    missing = [p for p in all_imgs if p.resolve() not in used]
    if unmatched or missing:
        raise RuntimeError(f'Imagens sem encaixe temático: unmatched={unmatched}; missing={[p.name for p in missing]}')

    # Group figures by the boundary paragraph. Insert in reverse order because
    # each XML insertion is made immediately before that same boundary.
    groups = defaultdict(list)
    for boundary, src, topic, target_title in placements:
        groups[id(boundary)].append((boundary, src, topic, target_title))
    for group in groups.values():
        boundary = group[0][0]
        for _, src, topic, target_title in reversed(group):
            add_figure_before(boundary, doc._body, optimized_image(src),
                              f'Figura — {topic} | Secção: {target_title}')

    doc.core_properties.title = 'Manual IPT — 4ª Edição (Livro Ilustrado)'
    doc.core_properties.subject = 'Manual de Inglês para Propósitos Técnicos com imagens didáticas'
    doc.core_properties.comments = '117 imagens inseridas no fim dos tópicos correspondentes.'
    doc.save(str(OUT))
    print(f'placements={len(placements)} total_images={len(all_imgs)} unmatched=0 missing=0')
    print(f'output={OUT}')


if __name__ == '__main__':
    main()
