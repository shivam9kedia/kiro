"""
Self-contained Markdown -> DOCX converter using only Python stdlib.
Built specifically for the Female-Audience-Hotspots-India-Geo-Targeting.md doc.

Handles: H1/H2/H3, paragraphs, bold/italic/inline-code, pipe tables,
bullet lists, ordered lists, code fences, horizontal rules, blockquotes.
"""

import os
import re
import zipfile
from xml.sax.saxutils import escape

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'


# ----------------------------- Inline run rendering -----------------------------

_DOC_HYPERLINKS = []  # accumulates (rId, url) during a single document build


def render_runs(text, default_rpr=""):
    """Render markdown inline formatting into Word XML runs.
    Supports [text](url) links, **bold**, *italic*, `code`, and plain text.
    """
    out = []
    # Tokenize. Link pattern MUST come first so it wins over * / ` patterns.
    pattern = re.compile(r'(\[[^\]]+\]\([^)]+\)|\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)')
    link_re = re.compile(r'^\[([^\]]+)\]\(([^)]+)\)$')
    parts = pattern.split(text)
    for part in parts:
        if not part:
            continue
        link_m = link_re.match(part)
        if link_m:
            link_text, url = link_m.group(1), link_m.group(2)
            rid = f'rIdHl{len(_DOC_HYPERLINKS) + 1}'
            _DOC_HYPERLINKS.append((rid, url))
            out.append(
                f'<w:hyperlink r:id="{rid}">'
                f'<w:r><w:rPr>{default_rpr}<w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr>'
                f'<w:t xml:space="preserve">{escape(link_text)}</w:t></w:r>'
                f'</w:hyperlink>'
            )
        elif part.startswith('**') and part.endswith('**'):
            content = part[2:-2]
            out.append(f'<w:r><w:rPr>{default_rpr}<w:b/></w:rPr><w:t xml:space="preserve">{escape(content)}</w:t></w:r>')
        elif part.startswith('*') and part.endswith('*') and len(part) > 2:
            content = part[1:-1]
            out.append(f'<w:r><w:rPr>{default_rpr}<w:i/></w:rPr><w:t xml:space="preserve">{escape(content)}</w:t></w:r>')
        elif part.startswith('`') and part.endswith('`'):
            content = part[1:-1]
            out.append(f'<w:r><w:rPr>{default_rpr}<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="20"/><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/></w:rPr><w:t xml:space="preserve">{escape(content)}</w:t></w:r>')
        else:
            out.append(f'<w:r><w:rPr>{default_rpr}</w:rPr><w:t xml:space="preserve">{escape(part)}</w:t></w:r>')
    return ''.join(out)


# ----------------------------- Block builders -----------------------------

def heading(text, level):
    style = f'Heading{level}'
    return f'<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr>{render_runs(text)}</w:p>'


def paragraph(text):
    return f'<w:p>{render_runs(text)}</w:p>'


def bullet(text, level=0):
    indent = 360 + (level * 360)
    return (
        f'<w:p><w:pPr><w:pStyle w:val="ListBullet"/>'
        f'<w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="1"/></w:numPr>'
        f'<w:ind w:left="{indent}" w:hanging="360"/></w:pPr>'
        f'{render_runs(text)}</w:p>'
    )


def numbered(text, level=0):
    indent = 360 + (level * 360)
    return (
        f'<w:p><w:pPr><w:pStyle w:val="ListNumber"/>'
        f'<w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="2"/></w:numPr>'
        f'<w:ind w:left="{indent}" w:hanging="360"/></w:pPr>'
        f'{render_runs(text)}</w:p>'
    )


def blockquote(text):
    quote_rpr = '<w:i/><w:color w:val="595959"/>'
    return (
        f'<w:p><w:pPr><w:pStyle w:val="Quote"/>'
        f'<w:ind w:left="720"/>'
        f'<w:pBdr><w:left w:val="single" w:sz="24" w:space="8" w:color="4472C4"/></w:pBdr>'
        f'</w:pPr>{render_runs(text, default_rpr=quote_rpr)}</w:p>'
    )


def horizontal_rule():
    return (
        '<w:p><w:pPr><w:pBdr>'
        '<w:bottom w:val="single" w:sz="6" w:space="1" w:color="BFBFBF"/>'
        '</w:pBdr></w:pPr></w:p>'
    )


def code_block(lines):
    body = ''
    for line in lines:
        body += (
            f'<w:p><w:pPr><w:pStyle w:val="Code"/>'
            f'<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="18"/></w:rPr>'
            f'<w:t xml:space="preserve">{escape(line)}</w:t></w:r></w:p>'
        )
    return body


def table(rows):
    """rows: list of list-of-strings. First row is header."""
    if not rows:
        return ''
    n_cols = len(rows[0])
    # Equal-width columns; total width ~= 9000 twips (page width ~ 9360 twips for A4 with std margins)
    col_w = 9000 // max(n_cols, 1)

    grid = ''.join(f'<w:gridCol w:w="{col_w}"/>' for _ in range(n_cols))

    tbl_pr = (
        '<w:tblPr>'
        '<w:tblW w:w="9000" w:type="dxa"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="D9D9D9"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="D9D9D9"/>'
        '</w:tblBorders>'
        '<w:tblLayout w:type="fixed"/>'
        '</w:tblPr>'
    )

    body = ''
    for i, row in enumerate(rows):
        is_header = (i == 0)
        shading = '<w:shd w:val="clear" w:color="auto" w:fill="4472C4"/>' if is_header else ''
        tr_props = '<w:trPr><w:tblHeader/></w:trPr>' if is_header else ''
        cells = ''
        for cell in row:
            cell_text = cell.strip()
            rpr = '<w:b/><w:color w:val="FFFFFF"/>' if is_header else ''
            cells += (
                f'<w:tc><w:tcPr><w:tcW w:w="{col_w}" w:type="dxa"/>{shading}</w:tcPr>'
                f'<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>'
                f'{render_runs(cell_text, default_rpr=rpr)}</w:p></w:tc>'
            )
        body += f'<w:tr>{tr_props}{cells}</w:tr>'

    # Add a small empty paragraph after the table so Word doesn't merge it with following content
    return f'<w:tbl>{tbl_pr}<w:tblGrid>{grid}</w:tblGrid>{body}</w:tbl><w:p/>'


# ----------------------------- Markdown parser -----------------------------

def parse_markdown(md):
    """Yield document XML chunks block by block."""
    lines = md.split('\n')
    i = 0
    n = len(lines)

    while i < n:
        line = lines[i]
        stripped = line.strip()

        # Skip blank lines
        if not stripped:
            i += 1
            continue

        # Code fence
        if stripped.startswith('```'):
            j = i + 1
            code_lines = []
            while j < n and not lines[j].strip().startswith('```'):
                code_lines.append(lines[j])
                j += 1
            yield code_block(code_lines)
            i = j + 1
            continue

        # Horizontal rule
        if re.match(r'^-{3,}$', stripped):
            yield horizontal_rule()
            i += 1
            continue

        # Headings
        m = re.match(r'^(#{1,6})\s+(.*)$', stripped)
        if m:
            level = min(len(m.group(1)), 4)  # cap at H4 since we only style up to H3
            if level > 3:
                level = 3
            yield heading(m.group(2).strip(), level)
            i += 1
            continue

        # Blockquote
        if stripped.startswith('>'):
            quote_lines = []
            while i < n and lines[i].strip().startswith('>'):
                quote_lines.append(re.sub(r'^>\s?', '', lines[i].strip()))
                i += 1
            yield blockquote(' '.join(quote_lines))
            continue

        # Pipe table — header line followed by separator |---|---|
        if '|' in stripped and i + 1 < n and re.match(r'^\s*\|?[\s\-:|]+\|?\s*$', lines[i + 1]):
            # collect table rows
            def split_row(s):
                s = s.strip()
                if s.startswith('|'):
                    s = s[1:]
                if s.endswith('|'):
                    s = s[:-1]
                return [c.strip() for c in s.split('|')]

            header = split_row(lines[i])
            i += 2  # skip header + separator
            rows = [header]
            while i < n and '|' in lines[i] and lines[i].strip():
                rows.append(split_row(lines[i]))
                i += 1
            yield table(rows)
            continue

        # Bullet list
        if re.match(r'^[\-\*]\s+', stripped):
            while i < n and re.match(r'^[\-\*]\s+', lines[i].strip()):
                content = re.sub(r'^[\-\*]\s+', '', lines[i].strip())
                yield bullet(content)
                i += 1
            continue

        # Numbered list
        if re.match(r'^\d+\.\s+', stripped):
            while i < n and re.match(r'^\d+\.\s+', lines[i].strip()):
                content = re.sub(r'^\d+\.\s+', '', lines[i].strip())
                yield numbered(content)
                i += 1
            continue

        # Paragraph (collect consecutive non-empty, non-special lines)
        para_lines = [stripped]
        i += 1
        while i < n:
            s = lines[i].strip()
            if not s:
                break
            # Stop if we hit a special block start
            if (s.startswith('#') or s.startswith('```') or s.startswith('>') or
                re.match(r'^-{3,}$', s) or re.match(r'^[\-\*]\s+', s) or
                re.match(r'^\d+\.\s+', s) or
                ('|' in s and i + 1 < n and re.match(r'^\s*\|?[\s\-:|]+\|?\s*$', lines[i + 1]))):
                break
            para_lines.append(s)
            i += 1
        yield paragraph(' '.join(para_lines))


# ----------------------------- DOCX assembly -----------------------------

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''

ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>'''

CORE_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Female Audience Hotspots - India Geo-Targeting</dc:title>
  <dc:creator>Marketing Research</dc:creator>
  <cp:lastModifiedBy>Marketing Research</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-05-22T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-05-22T00:00:00Z</dcterms:modified>
</cp:coreProperties>'''

APP_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Custom DOCX Generator</Application>
</Properties>'''

STYLES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:lang w:val="en-IN"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="120" w:line="276" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:spacing w:before="360" w:after="160"/>
      <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
      <w:b/>
      <w:color w:val="1F3864"/>
      <w:sz w:val="40"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:spacing w:before="280" w:after="120"/>
      <w:outlineLvl w:val="1"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
      <w:b/>
      <w:color w:val="2E74B5"/>
      <w:sz w:val="30"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:spacing w:before="200" w:after="100"/>
      <w:outlineLvl w:val="2"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
      <w:b/>
      <w:color w:val="4472C4"/>
      <w:sz w:val="26"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListBullet">
    <w:name w:val="List Bullet"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:after="60"/>
    </w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListNumber">
    <w:name w:val="List Number"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:after="60"/>
    </w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Quote">
    <w:name w:val="Quote"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="120" w:after="120"/>
    </w:pPr>
    <w:rPr>
      <w:i/>
      <w:color w:val="595959"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Code">
    <w:name w:val="Code"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>
      <w:sz w:val="18"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="0" w:after="240"/>
      <w:jc w:val="center"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
      <w:b/>
      <w:color w:val="1F3864"/>
      <w:sz w:val="48"/>
    </w:rPr>
  </w:style>
</w:styles>'''

NUMBERING_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#8226;"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="o"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>'''


def build_document_xml(md_text):
    body_chunks = list(parse_markdown(md_text))
    body = ''.join(body_chunks)

    sect_pr = (
        '<w:sectPr>'
        '<w:pgSz w:w="11906" w:h="16838"/>'  # A4
        '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" '
        'w:header="720" w:footer="720" w:gutter="0"/>'
        '</w:sectPr>'
    )

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document {W_NS} {R_NS}>'
        f'<w:body>{body}{sect_pr}</w:body>'
        '</w:document>'
    )


def build_docx(md_path, out_path):
    global _DOC_HYPERLINKS
    _DOC_HYPERLINKS = []  # reset per document

    with open(md_path, 'r', encoding='utf-8') as f:
        md_text = f.read()

    document_xml = build_document_xml(md_text)  # populates _DOC_HYPERLINKS

    # Build document rels dynamically: styles + numbering + one per hyperlink.
    hyperlink_rels = ''.join(
        f'<Relationship Id="{rid}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
        f'Target="{escape(url)}" TargetMode="External"/>'
        for rid, url in _DOC_HYPERLINKS
    )
    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
        f'{hyperlink_rels}'
        '</Relationships>'
    )

    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', doc_rels)
        z.writestr('word/document.xml', document_xml)
        z.writestr('word/styles.xml', STYLES_XML)
        z.writestr('word/numbering.xml', NUMBERING_XML)
        z.writestr('docProps/core.xml', CORE_XML)
        z.writestr('docProps/app.xml', APP_XML)

    print(f'Wrote {out_path} ({os.path.getsize(out_path):,} bytes, {len(_DOC_HYPERLINKS)} hyperlinks)')


if __name__ == '__main__':
    here = os.path.dirname(os.path.abspath(__file__))
    md = os.path.join(here, 'Female-Audience-Hotspots-India-Geo-Targeting.md')
    out = os.path.join(here, 'Female-Audience-Hotspots-India-Geo-Targeting.docx')
    build_docx(md, out)
