"""Собрать рабочий DOCX из разделов курсовой. Поля обновляются в Word."""
import re
from itertools import groupby
from pathlib import Path
from tempfile import NamedTemporaryFile
from os import replace
from xml.sax.saxutils import escape, quoteattr
from zipfile import ZipFile, ZIP_DEFLATED

ROOT = Path(__file__).resolve().parents[1]
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
FONT = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
SECTIONS = ['chapters/abstract.md', 'chapters/glossary.md', 'chapters/00-introduction.md',
            'chapters/01-theoretical.md', 'chapters/02-empirical.md', 'chapters/03-conclusion.md',
            'chapters/04-bibliography.md', 'appendices/01-comparison.md']


def run(text, bold=False):
    fmt = '<w:rPr><w:b/></w:rPr>' if bold else ''
    return f'<w:r>{fmt}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def paragraph(content, style='Normal'):
    return f'<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr>{content}</w:p>'


def style(name, fmt, size=14, bold=False):
    base = '' if name == 'Normal' else '<w:basedOn w:val="Normal"/>'
    return (f'<w:style w:type="paragraph" w:styleId="{name}"><w:name w:val="{name}"/>{base}'
            f'<w:pPr>{fmt}</w:pPr><w:rPr>{FONT}{"<w:b/>" if bold else ""}'
            f'<w:sz w:val="{2*size}"/><w:lang w:val="ru-RU"/></w:rPr></w:style>')


def field(instruction):
    return (f'<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>'
            f'<w:r><w:instrText xml:space="preserve"> {instruction} </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>')


def relationships(items):
    entries = ''.join(f'<Relationship Id="{ident}" Type="{R}/{kind}" Target={quoteattr(target)}'
                      + (' TargetMode="External"' if kind == 'hyperlink' else '') + '/>'
                      for ident, kind, target in items)
    return f'<Relationships xmlns="{PKG}">{entries}</Relationships>'


def render_text(line, links):
    parts = []
    for token in re.split(r'(\*\*.*?\*\*|\[[^\]]+\]\(https?://[^)]+\))', line):
        match = re.fullmatch(r'\[([^\]]+)\]\((https?://[^)]+)\)', token)
        if match:
            label, url = match.groups()
            ident = links.setdefault(url, f'link{len(links)+1}')
            # При печати локатор первоисточника остаётся видимым.
            label = url if label.startswith('Полный') else '[' + label + ']'
            parts.append(f'<w:hyperlink r:id="{ident}">{run(label)}</w:hyperlink>')
        elif token:
            bold = token.startswith('**')
            parts.append(run(token[2:-2] if bold else token, bold))
    return ''.join(parts)


def table(lines):
    rows = [[c.strip() for c in line.strip('|').split('|')] for line in lines]
    del rows[1]  # разделитель Markdown
    width = 9638 // len(rows[0])
    grid = f'<w:gridCol w:w="{width}"/>' * len(rows[0])
    borders = ''.join(f'<w:{edge} w:val="single" w:sz="4"/>' for edge in
                      ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'])
    out = [f'<w:tbl><w:tblPr><w:tblW w:w="9638" w:type="dxa"/><w:tblBorders>{borders}</w:tblBorders>'
           f'</w:tblPr><w:tblGrid>{grid}</w:tblGrid>']
    for i, row in enumerate(rows):
        out.append('<w:tr><w:trPr><w:cantSplit/>' + ('<w:tblHeader/>' if i == 0 else '') + '</w:trPr>')
        for value in row:
            cell = paragraph(run(value, i == 0), 'Cell' if i == len(rows)-1 else 'TableKeep')
            out.append(f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>{cell}</w:tc>')
        out.append('</w:tr>')
    return ''.join(out) + '</w:tbl>'


def build():
    title = dict(re.findall(r'^\| ([^|]+?) \| ([^|]+?) \|$', (ROOT/'chapters/title-page.md').read_text(), re.M))
    out = [paragraph(run('________________________________________'), 'Center'),
           paragraph(run('Факультет и кафедра ________________________'), 'Center'),
           paragraph(run(title['Вид работы'].upper(), True), 'Title'),
           paragraph(run('по дисциплине ________________________'), 'Center'),
           paragraph(run('Тема «' + title['Тема'] + '»', True), 'Center'),
           paragraph(run('Выполнил(а) ____________________, 3 курс, группа ______'), 'TitleLine'),
           paragraph(run('Руководитель _________________________________'), 'Center'),
           paragraph(run('____________________, 20___ г.'), 'TitleLine'),
           paragraph(run('Содержание'), 'Contents'), paragraph(field(r'TOC \o "1-2" \h \z \u'), 'Cell')]
    links = {}
    for filename in SECTIONS:
        source = (ROOT / filename).read_text()
        source = re.sub(r'<!--.*?-->', '', source, flags=re.S)
        for is_table, group in groupby(source.splitlines(), lambda line: line.startswith('|')):
            if is_table:
                out.append(table(group))
                continue
            for line in filter(None, group):
                if line.startswith('# '):
                    out.append(paragraph(run(line[2:]), 'Heading1'))
                elif line.startswith('## '):
                    out.append(paragraph(run(line[3:]), 'Heading2'))
                else:
                    fmt = 'Caption' if re.fullmatch(r'\*\*.+\*\*', line) else 'Normal'
                    out.append(paragraph(render_text(line, links), fmt))
    section = ('<w:sectPr><w:headerReference w:type="default" r:id="header"/><w:pgSz w:w="11906" w:h="16838"/>'
               '<w:pgMar w:top="1134" w:right="567" w:bottom="1134" w:left="1701" w:header="567" w:footer="567"/>'
               '<w:titlePg/></w:sectPr>')
    root = f'xmlns:w="{W}" xmlns:r="{R}"'
    indent = '<w:ind w:firstLine="0"/>'
    center = indent + '<w:jc w:val="center"/>'
    major = '<w:keepNext/><w:pageBreakBefore/><w:spacing w:after="480"/>' + center
    fmt = [style('Normal', '<w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/>'
                 '<w:ind w:firstLine="709"/><w:jc w:val="both"/>'),
           style('Heading1', major + '<w:outlineLvl w:val="0"/>', 16, True),
           style('Heading2', '<w:keepNext/><w:spacing w:before="360" w:after="240"/>' + indent +
                 '<w:outlineLvl w:val="1"/>', 14, True), style('Contents', major, 16, True),
           style('Center', center), style('Title', '<w:spacing w:before="2000"/>' + center, 16, True),
           style('TitleLine', '<w:spacing w:before="1200"/>' + center),
           style('Cell', indent + '<w:jc w:val="left"/>'), style('TableKeep', '<w:keepNext/>' + indent + '<w:jc w:val="left"/>'), style('Caption', '<w:keepNext/>' + indent)]
    for level in (1, 2):
        fmt.append(style(f'TOC{level}', f'<w:spacing w:line="240" w:after="0" w:before="0"/>'
                         f'<w:ind w:left="{(level-1)*180}" w:firstLine="0"/><w:jc w:val="left"/>'))
    rels = [('styles', 'styles', 'styles.xml'), ('header', 'header', 'header1.xml'),
            ('settings', 'settings', 'settings.xml')] + [(v, 'hyperlink', k) for k, v in links.items()]
    mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.'
    parts = {'document': ('document.main', f'<w:document {root}><w:body>{"".join(out)}{section}</w:body></w:document>'),
             'header1': ('header', f'<w:hdr {root}><w:p><w:pPr><w:jc w:val="right"/></w:pPr>{field("PAGE")}</w:p></w:hdr>'),
             'styles': ('styles', f'<w:styles {root}>{"".join(fmt)}</w:styles>'),
             'settings': ('settings', f'<w:settings {root}><w:updateFields w:val="true"/></w:settings>')}
    types = ''.join(f'<Override PartName="/word/{k}.xml" ContentType="{mime}{v[0]}+xml"/>' for k,v in parts.items())
    dest = ROOT / 'output/coursework.docx'
    with NamedTemporaryFile(dir=dest.parent, suffix='.tmp', delete=False) as temp:
        tmp_path = Path(temp.name)
    try:
        with ZipFile(tmp_path, 'w', ZIP_DEFLATED) as doc:
            doc.writestr('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                         '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                         f'{types}</Types>')
            doc.writestr('_rels/.rels', relationships([('doc', 'officeDocument', 'word/document.xml')]))
            doc.writestr('word/_rels/document.xml.rels', relationships(rels))
            for key, (_, data) in parts.items():
                doc.writestr('word/' + key + '.xml', data)
        replace(tmp_path, dest)
    finally:
        tmp_path.unlink(missing_ok=True)


if __name__ == '__main__':
    build()
