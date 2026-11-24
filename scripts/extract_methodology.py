#!/usr/bin/env python3
"""Extract the supplied DOCX into sectioned Markdown using only the standard library."""

from pathlib import Path
import re
from xml.etree import ElementTree as ET
from zipfile import ZipFile


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / 'docs/Методичка_по_выполнению_курсовой_работы_6.docx'
DEST = ROOT / 'requirements/methodology'
NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


def paragraph_text(node):
    parts = []
    for child in node.iter():
        if child.tag == f"{{{NS['w']}}}t":
            parts.append(child.text or '')
        elif child.tag == f"{{{NS['w']}}}tab":
            parts.append('\t')
        elif child.tag in (f"{{{NS['w']}}}br", f"{{{NS['w']}}}cr"):
            parts.append('\n')
    return ''.join(parts).strip()


def main():
    with ZipFile(SOURCE) as archive:
        body = ET.fromstring(archive.read('word/document.xml')).find('w:body', NS)
    sections = []
    next_section = 1
    title, filename, blocks = 'Титульные сведения методички', '00-title.md', []
    for node in body:
        if node.tag == f"{{{NS['w']}}}p":
            text = paragraph_text(node)
            if not text:
                continue
            section = re.match(r'^(\d{1,2})\.\s+\S', text)
            appendix = re.match(r'^Приложение №(\d+(?:\.\d+)?)\.?\s+(?=Пример|Образец|Оформление)', text)
            # Numbered lists elsewhere must not become section boundaries.
            heading = section and int(section[1]) == next_section
            if appendix or heading:
                sections.append((title, filename, blocks))
                title = text
                filename = f'appendix-{appendix[1]}.md' if appendix else f'{int(section[1]):02d}-section.md'
                blocks = []
                if heading:
                    next_section += 1
            else:
                blocks.append(text)
        elif node.tag == f"{{{NS['w']}}}tbl":
            rows = []
            for row in node.findall('w:tr', NS):
                cells = []
                for cell in row.findall('w:tc', NS):
                    value = '<br>'.join(filter(None, (paragraph_text(p) for p in cell.findall('w:p', NS))))
                    cells.append(value.replace('|', '\\|').replace('\n', '<br>'))
                rows.append('| ' + ' | '.join(cells) + ' |')
            if rows:
                rows.insert(1, '| ' + ' | '.join(['---'] * len(node.findall('w:tr', NS)[0].findall('w:tc', NS))) + ' |')
                blocks.append('\n'.join(rows))
    sections.append((title, filename, blocks))
    DEST.mkdir(parents=True, exist_ok=True)
    links = []
    for title, filename, blocks in sections:
        (DEST / filename).write_text(f'# {title}\n\n' + '\n\n'.join(blocks) + '\n', encoding='utf-8')
        links.append(f'- [{title}]({filename})')
    (DEST / 'README.md').write_text(
        '# Методические рекомендации: извлечение из DOCX\n\n'
        'Источник: [оригинал](../../docs/Методичка_по_выполнению_курсовой_работы_6.docx), '
        'ОЧУ ВО «Московская международная академия», Москва, 2019.\n\n'
        'Полный текст основного тела документа разделён по исходным разделам и приложениям. '
        'Таблица глоссария сохранена. Орфография, примеры, повторы и противоречия оригинала не исправлялись. '
        'Вёрстка, колонтитулы, автоматические номера страниц и пустые абзацы не воспроизводятся. '
        'Это текстовая копия для навигации; внешний вид проверяется по DOCX.\n\n'
        'Повторное извлечение: `python3 scripts/extract_methodology.py` из корня проекта. '
        'Генерируемые файлы этого каталога вручную не редактировать.\n\n'
        + '\n'.join(links) + '\n', encoding='utf-8')
    print(f'Extracted {len(sections)} sections into {DEST.relative_to(ROOT)}')


if __name__ == '__main__':
    main()
