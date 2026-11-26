"""ДЛЯ ТЕСТОВ. Читает принятые множители из фактически сохранённого Results без запуска Excel."""

from argparse import ArgumentParser
import json
from pathlib import Path, PurePosixPath
import re
from xml.etree import ElementTree as ET
from zipfile import ZipFile


def TestSavedTrace(root, report_name='Stock'):
    """Проверяет нативное сохранение/открытие и возвращает точные числовые записи трассы из XML листа Results."""
    report = (root / ('Native/' + report_name + '.txt')).read_text('utf-8-sig')
    matches = re.findall(r'^RESULTS_SAVE_REOPEN: .*equal=True; file=([^\r\n]+)$', report, re.M)
    assert len(matches) == 1 and 'WATCHDOG_COMPLETED: exit=0' in report
    saved = Path(matches[0])
    markup = root / 'SavedWorkbookMarkup.zip'
    if markup.exists():
        scope = json.loads((root / 'SavedWorkbookScope.json').read_bytes())
        scratch = Path(scope['original_scratch_root'])
        assert saved.resolve().is_relative_to((scratch / 'Temp').resolve())
        saved = markup
    else:
        assert saved.resolve().is_relative_to((root / 'Temp').resolve())
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    with ZipFile(saved) as book:
        strings = []
        if 'xl/sharedStrings.xml' in book.namelist():
            strings = [''.join(item.itertext()) for item in ET.fromstring(book.read('xl/sharedStrings.xml'))]
        sheets = ET.fromstring(book.read('xl/workbook.xml')).find('s:sheets', ns)
        sheet, = [item for item in sheets if item.attrib['name'] == 'Results']
        relation = sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
        relationships = ET.fromstring(book.read('xl/_rels/workbook.xml.rels'))
        target, = [item.attrib['Target'] for item in relationships if item.attrib['Id'] == relation]
        name = target.lstrip('/') if target.startswith('/') else str(PurePosixPath('xl') / target)
        rows = ET.fromstring(book.read(name)).find('s:sheetData', ns)
        result = {}
        for row in rows:
            cells = []
            for cell in row:
                value = cell.find('s:v', ns)
                text = '' if value is None else value.text
                if cell.attrib.get('t') == 's':
                    text = strings[int(text)]
                elif cell.attrib.get('t') == 'inlineStr':
                    text = ''.join(cell.find('s:is', ns).itertext())
                cells.append((cell.attrib['r'], text))
            for i, (address, text) in enumerate(cells):
                if re.fullmatch(r'Coupled\.Trace\.(?:Direct|Capacity)\.\d+\.(?:Lambda|Offset)', text or ''):
                    assert text not in result and i + 1 < len(cells)
                    result[text] = {'address': address, 'value': cells[i + 1][1]}
    assert result, 'No exact coupled trace keys found'
    return result


def TestCompare(before, after, output):
    """Сравнивает порядок принятых точек без округления; различие не объявляется физической ошибкой автоматически."""
    old, new = TestSavedTrace(before), TestSavedTrace(after)
    prefix = 0
    while True:
        index = prefix + 1
        keys = [f'Coupled.Trace.Capacity.{index}.{kind}' for kind in ('Offset', 'Lambda')]
        if not all(key in old and key in new for key in keys):
            break
        if any(old[key]['value'] != new[key]['value'] for key in keys):
            break
        prefix = index
    result = {'before': old, 'after': new, 'capacity_trace_exact_common_prefix': prefix,
              'full_field_prefix_not_verified_by_load_trace_alone': True, 'physical_capacity_certified': False}
    output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print('SAVED RESULTS TRACE EXACT COMMON PREFIX', prefix)
    for root, data in ((before, old), (after, new)):
        loads = [float(item['value']) for key, item in data.items() if '.Capacity.' in key and key.endswith('.Lambda')]
        print(root.name, 'capacity states', len(loads), 'max lambda', max(loads))


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--before', type=Path, required=True)
    parser.add_argument('--after', type=Path, required=True)
    parser.add_argument('--out', type=Path, required=True)
    args = parser.parse_args()
    TestCompare(args.before, args.after, args.out)
