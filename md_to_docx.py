#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Markdown -> DOCX konvertor (faqat Python standart kutubxonasi).
BMI hujjatini Word formatiga aylantiradi: Times New Roman 14pt, 1.5 interval,
sarlavhalar, jadvallar, kod bloklari, diagramma bloklari, ro'yxatlar.
"""
import re
import os
import zipfile
from xml.sax.saxutils import escape as _xe

SRC = "BMI_Intellektual_Bookcrossing_Platformasi.md"
OUT = "BMI_Intellektual_Bookcrossing_Platformasi.docx"

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def esc(s):
    return _xe(s, {'"': "&quot;"})


# ---------- Inline parsing: **bold**, *italic*, `code` ----------
TOKEN_RE = re.compile(r'(\*\*.+?\*\*|\*[^*]+?\*|`[^`]+?`)')


def runs_from_text(text, base_sz=28, base_font="Times New Roman"):
    """Matnni run (w:r) larga aylantiradi, ichki **bold**/*italic*/`code` bilan."""
    # escaped underscores/hashes
    text = text.replace('\\_', '_').replace('\\#', '#').replace('\\*', '\u0001')
    out = []
    parts = TOKEN_RE.split(text)
    for part in parts:
        if not part:
            continue
        bold = italic = code = False
        t = part
        if part.startswith('**') and part.endswith('**') and len(part) >= 4:
            bold = True
            t = part[2:-2]
        elif part.startswith('*') and part.endswith('*') and len(part) >= 2:
            italic = True
            t = part[1:-1]
        elif part.startswith('`') and part.endswith('`') and len(part) >= 2:
            code = True
            t = part[1:-1]
        t = t.replace('\u0001', '*')
        rpr = []
        font = "Courier New" if code else base_font
        rpr.append('<w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>' % (font, font, font))
        if bold:
            rpr.append('<w:b/>')
        if italic:
            rpr.append('<w:i/>')
        sz = 22 if code else base_sz
        rpr.append('<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (sz, sz))
        if code:
            rpr.append('<w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/>')
        run = ('<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r>'
               % (''.join(rpr), esc(t)))
        out.append(run)
    return ''.join(out) if out else (
        '<w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"/>'
        '<w:sz w:val="%d"/></w:rPr><w:t xml:space="preserve"></w:t></w:r>'
        % (base_font, base_font, base_sz))


def para(text, *, sz=28, bold=False, italic=False, align='both', spacing_line=360,
         before=0, after=120, page_break=False, center=False, ind_left=0,
         keep_next=False, color=None):
    if center:
        align = 'center'
    ppr = ['<w:pStyle w:val="Normal"/>']
    pr_inner = []
    if page_break:
        pr_inner.append('<w:pageBreakBefore/>')
    if keep_next:
        pr_inner.append('<w:keepNext/>')
    pr_inner.append('<w:spacing w:line="%d" w:lineRule="auto" w:before="%d" w:after="%d"/>'
                    % (spacing_line, before, after))
    pr_inner.append('<w:jc w:val="%s"/>' % align)
    if ind_left:
        pr_inner.append('<w:ind w:left="%d"/>' % ind_left)
    # paragraph-level run props (affect whole para if no inner override)
    rpr = ['<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>',
           '<w:sz w:val="%d"/>' % sz]
    if bold:
        rpr.append('<w:b/>')
    if italic:
        rpr.append('<w:i/>')
    if color:
        rpr.append('<w:color w:val="%s"/>' % color)
    ppr_xml = '<w:pPr>%s<w:rPr>%s</w:rPr></w:pPr>' % (''.join(pr_inner), ''.join(rpr))

    # build runs
    if bold or italic or color:
        rpr2 = ['<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>',
                '<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (sz, sz)]
        if bold:
            rpr2.append('<w:b/>')
        if italic:
            rpr2.append('<w:i/>')
        if color:
            rpr2.append('<w:color w:val="%s"/>' % color)
        runs = ('<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r>'
                % (''.join(rpr2), esc(text.replace('\\_', '_'))))
    else:
        runs = runs_from_text(text, base_sz=sz)
    return '<w:p>%s%s</w:p>' % (ppr_xml, runs)


def empty_para(after=0):
    return ('<w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto" w:after="%d"/>'
            '<w:rPr><w:sz w:val="20"/></w:rPr></w:pPr></w:p>' % after)


def code_block(lines, mermaid=False):
    """Kod yoki diagramma blokini ramkali, soyali bitta katakli jadval sifatida."""
    fill = "F2F7FF" if mermaid else "F4F4F4"
    bcolor = "5B8DEF" if mermaid else "BBBBBB"
    cell_paras = []
    if mermaid:
        cell_paras.append(
            '<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto" w:after="40"/>'
            '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            '<w:b/><w:i/><w:sz w:val="18"/><w:color w:val="2E5BBB"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            '<w:b/><w:i/><w:sz w:val="18"/><w:color w:val="2E5BBB"/></w:rPr>'
            '<w:t xml:space="preserve">[ Diagramma — quyidagi sxema asosida chizilgan (Mermaid manbasi) ]</w:t></w:r></w:p>')
    for ln in lines:
        ln = ln.replace('\t', '    ')
        cell_paras.append(
            '<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto" w:after="0"/>'
            '<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/>'
            '<w:sz w:val="16"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/>'
            '<w:sz w:val="16"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % esc(ln))
    if not cell_paras:
        cell_paras.append(empty_para())
    tbl = (
        '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="8" w:color="%(b)s"/>'
        '<w:left w:val="single" w:sz="8" w:color="%(b)s"/>'
        '<w:bottom w:val="single" w:sz="8" w:color="%(b)s"/>'
        '<w:right w:val="single" w:sz="8" w:color="%(b)s"/>'
        '</w:tblBorders>'
        '<w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>'
        '<w:bottom w:w="60" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar>'
        '</w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="9600"/></w:tblGrid>'
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="pct"/>'
        '<w:shd w:val="clear" w:color="auto" w:fill="%(f)s"/></w:tcPr>%(c)s</w:tc></w:tr>'
        '</w:tbl>' % {'b': bcolor, 'f': fill, 'c': ''.join(cell_paras)})
    return tbl + empty_para(after=120)


def make_table(rows):
    """rows: list of list of cell-strings. Birinchi qator — sarlavha."""
    ncols = max(len(r) for r in rows)
    colw = int(9600 / ncols)
    grid = ''.join('<w:gridCol w:w="%d"/>' % colw for _ in range(ncols))
    trs = []
    for ri, row in enumerate(rows):
        header = (ri == 0)
        cells = []
        # pad
        row = row + [''] * (ncols - len(row))
        for cell in row:
            shd = ('<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/>'
                   if header else '')
            align = 'center' if header else 'both'
            rpr = ('<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
                   '<w:sz w:val="22"/><w:szCs w:val="22"/>')
            if header:
                rpr += '<w:b/>'
            runs = runs_from_text(cell, base_sz=22)
            if header:
                # force bold by wrapping: simpler — rebuild as bold text
                runs = ('<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r>'
                        % (rpr, esc(cell.replace('**', '').replace('\\_', '_'))))
            p = ('<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto" w:after="20"/>'
                 '<w:jc w:val="%s"/><w:rPr>%s</w:rPr></w:pPr>%s</w:p>'
                 % (align, rpr, runs))
            cells.append(
                '<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s'
                '<w:vAlign w:val="center"/></w:tcPr>%s</w:tc>' % (colw, shd, p))
        trs.append('<w:tr>%s%s</w:tr>' % (
            '<w:trPr><w:tblHeader/></w:trPr>' if header else '', ''.join(cells)))
    tbl = (
        '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/>'
        '<w:jc w:val="center"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="6" w:color="666666"/>'
        '<w:left w:val="single" w:sz="6" w:color="666666"/>'
        '<w:bottom w:val="single" w:sz="6" w:color="666666"/>'
        '<w:right w:val="single" w:sz="6" w:color="666666"/>'
        '<w:insideH w:val="single" w:sz="6" w:color="999999"/>'
        '<w:insideV w:val="single" w:sz="6" w:color="999999"/>'
        '</w:tblBorders>'
        '<w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="100" w:type="dxa"/>'
        '<w:bottom w:w="40" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar>'
        '</w:tblPr><w:tblGrid>%s</w:tblGrid>%s</w:tbl>' % (grid, ''.join(trs)))
    return tbl + empty_para(after=120)


def split_table_row(line):
    line = line.strip()
    if line.startswith('|'):
        line = line[1:]
    if line.endswith('|'):
        line = line[:-1]
    return [c.strip() for c in line.split('|')]


def is_sep_row(line):
    s = line.replace('|', '').replace('-', '').replace(':', '').strip()
    return s == '' and '-' in line


# ---------- Main parse ----------
def main():
    with open(SRC, encoding='utf-8') as f:
        text = f.read()
    lines = text.split('\n')
    body = []
    i = 0
    n = len(lines)
    first_h1_seen = False

    # caption detector
    cap_re = re.compile(r'^\*\*\d+\.\d+-rasm\.')
    is_fig_caption = lambda s: bool(cap_re.match(s.strip()))

    while i < n:
        line = lines[i]
        stripped = line.strip()

        # blank
        if stripped == '':
            i += 1
            continue

        # horizontal rule -> skip (used as separator)
        if stripped == '---':
            i += 1
            continue

        # fenced code block
        if stripped.startswith('```'):
            lang = stripped[3:].strip().lower()
            mermaid = (lang == 'mermaid' or lang.startswith('xychart')
                       or lang.startswith('%%') or lang in (
                           'erdiagram', 'sequencediagram', 'gantt', 'pie',
                           'quadrantchart', 'mindmap', 'flowchart'))
            code_lines = []
            i += 1
            while i < n and not lines[i].strip().startswith('```'):
                code_lines.append(lines[i])
                i += 1
            i += 1  # skip closing ```
            # detect mermaid by content if lang empty
            if not lang and code_lines:
                head = code_lines[0].strip().lower()
                if any(head.startswith(k) for k in (
                        'flowchart', 'graph', 'sequencediagram', 'erdiagram',
                        'gantt', 'pie', 'mindmap', 'xychart', 'quadrantchart',
                        '%%{')):
                    mermaid = True
            body.append(code_block(code_lines, mermaid=mermaid))
            continue

        # table
        if stripped.startswith('|') and i + 1 < n and is_sep_row(lines[i + 1]):
            rows = [split_table_row(line)]
            i += 1
            # skip separator
            i += 1
            while i < n and lines[i].strip().startswith('|'):
                if is_sep_row(lines[i]):
                    i += 1
                    continue
                rows.append(split_table_row(lines[i]))
                i += 1
            body.append(make_table(rows))
            continue

        # headings
        m = re.match(r'^(#{1,6})\s+(.*)$', stripped)
        if m:
            level = len(m.group(1))
            htext = m.group(2).strip()
            if level == 1:
                pb = first_h1_seen
                first_h1_seen = True
                body.append(para(htext, sz=32, bold=True, center=True,
                                 page_break=pb, before=120, after=240,
                                 align='center', keep_next=True))
            elif level == 2:
                body.append(para(htext, sz=28, bold=True, align='left',
                                 before=240, after=120, keep_next=True))
            elif level == 3:
                body.append(para(htext, sz=26, bold=True, italic=True,
                                 align='left', before=180, after=100, keep_next=True))
            else:
                body.append(para(htext, sz=24, bold=True, align='left',
                                 before=120, after=80, keep_next=True))
            i += 1
            continue

        # figure caption (bold line starting with N.N-rasm)
        if is_fig_caption(stripped):
            txt = stripped.strip('*').strip()
            body.append(para(txt, sz=24, bold=True, center=True,
                             before=120, after=60, keep_next=True))
            i += 1
            continue

        # list items (unordered)
        mu = re.match(r'^[-*]\s+(.*)$', stripped)
        if mu:
            body.append(para('•  ' + mu.group(1), sz=28, align='both',
                             ind_left=360, after=60))
            i += 1
            continue

        # ordered list
        mo = re.match(r'^(\d+)\.\s+(.*)$', stripped)
        if mo:
            body.append(para(mo.group(1) + '.  ' + mo.group(2), sz=28,
                             align='both', ind_left=360, after=60))
            i += 1
            continue

        # izoh (italic note line)
        if stripped.startswith('*') and stripped.endswith('*') and not stripped.startswith('**'):
            body.append(para(stripped.strip('*'), sz=24, italic=True,
                             align='both', after=120, color="444444"))
            i += 1
            continue

        # normal paragraph
        body.append(para(stripped, sz=28, align='both', after=120))
        i += 1

    # ---------- assemble document.xml ----------
    sect = (
        '<w:sectPr>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" '
        'w:header="708" w:footer="708" w:gutter="0"/>'
        '<w:pgNumType w:start="1"/>'
        '</w:sectPr>')

    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="%s">'
        '<w:body>%s%s</w:body></w:document>' % (W, ''.join(body), sect))

    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:styles xmlns:w="%s">'
        '<w:docDefaults><w:rPrDefault><w:rPr>'
        '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        '<w:sz w:val="28"/><w:szCs w:val="28"/><w:lang w:val="en-US"/>'
        '</w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr>'
        '<w:spacing w:line="360" w:lineRule="auto" w:after="120"/>'
        '</w:pPr></w:pPrDefault></w:docDefaults>'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/><w:qFormat/></w:style>'
        '</w:styles>' % W)

    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '</Types>')

    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')

    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '</Relationships>')

    if os.path.exists(OUT):
        os.remove(OUT)
    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', content_types)
        z.writestr('_rels/.rels', rels)
        z.writestr('word/document.xml', document)
        z.writestr('word/styles.xml', styles)
        z.writestr('word/_rels/document.xml.rels', doc_rels)

    print("DOCX yaratildi:", OUT)
    print("Body bloklar soni:", len(body))


if __name__ == '__main__':
    main()
