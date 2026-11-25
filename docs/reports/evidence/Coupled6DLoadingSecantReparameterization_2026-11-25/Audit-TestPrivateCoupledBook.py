"""ДЛЯ ТЕСТОВ. Сопоставляет все фактические VBA-компоненты изолированной книги с закреплёнными входами."""

from argparse import ArgumentParser
from difflib import unified_diff
from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path
import re
from runpy import run_path
import tarfile
from zipfile import ZipFile


def TestAudit(repo, root, output):
    """Неизменённые компоненты сравнивает с установленной базой; импортированные — с точными исходниками подготовки."""
    here = Path(__file__).resolve().parent
    prep = json.loads((root / 'Preparation.json').read_bytes())
    compilation = json.loads((root / 'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
    assert compilation['Compiled'] and not compilation['ProductionChanged']
    assert json.loads((root / 'CompileClosure.json').read_bytes())['State'] == 'Absent'
    source = (repo / 'workbook/output/RC_Section_NDM.xlsm').read_bytes()
    for name, expected in prep['inputs'].items():
        raw = source if name == 'workbook/output/RC_Section_NDM.xlsm' else (root / name).read_bytes()
        assert sha256(raw).hexdigest().upper() == expected, name
    assert sha256(source).hexdigest().upper() == prep['source_book_sha256'] == compilation['ProductionSHA256']
    helper, body, reader = run_path(str(here / 'Inspect-TestResolutionVBA.py'))['TestHelpers'](repo)
    canonical = run_path(str(here / 'Canonical-TestCoreVBA.py'))
    normalize = canonical['TestCanonical']
    wrapped = canonical['TestWrap'](body)
    _, original = reader(BytesIO(source))
    storage = 'workbook'
    if (root / 'CompiledVBA.solid.json').exists():
        solid = json.loads((root / 'CompiledVBA.solid.json').read_bytes())
        assert solid['format'] == 'NDM_TEST_SOLID_CFB_01'
        assert solid['bundle'] == 'CompiledProjects.tar.xz' and solid['member'] == root.name + '.bin'
        assert root.parent.name == 'Cases'
        packed = root.parent.parent / solid['bundle']
        assert sha256(packed.read_bytes()).hexdigest().upper() == solid['bundle_sha256']
        binary = None
        with tarfile.open(packed, mode='r|xz') as bundle:
            for member in bundle:
                assert member.isfile() and '/' not in member.name and member.size <= 64 * 1024 * 1024
                if member.name == solid['member']:
                    binary = bundle.extractfile(member).read()
                    break
        pin = json.loads((root / 'CompiledVBA.json').read_bytes())
        assert binary is not None and len(binary) == pin['bytes'] == solid['bytes']
        assert sha256(binary).hexdigest().upper() == pin['sha256'] == solid['sha256']
        storage = 'solid-exact-CFB'
    elif (root / 'CompiledVBA.bin').exists():
        binary = (root / 'CompiledVBA.bin').read_bytes()
        pin = json.loads((root / 'CompiledVBA.json').read_bytes())
        assert len(binary) == pin['bytes'] and sha256(binary).hexdigest().upper() == pin['sha256']
        storage = 'full-CFB'
    else:
        compiled = (root / 'docs/regression/Diagnostic.xlsm').read_bytes()
        assert sha256(compiled).hexdigest().upper() == compilation['WorkbookSHA256']
        with ZipFile(BytesIO(compiled)) as book:
            binary = book.read('xl/vbaProject.bin')
    _, actual = reader(helper['TestContainer'](binary))
    assert set(actual) == set(original) | set(compilation['Modules'])
    changes = {}
    unchanged = []
    def TestMatch(name, path):
        """Применяет один и тот же строгий контракт VBE к импортам и прочим сохранённым рабочим исходникам."""
        alias = 'modTestMVT2GaussSpan' if name == 'modTestMVT2ProductionASDTest' else name
        expected, edits = helper['TestMVTCompilerBody'](body, path.read_text('utf-8-sig'), alias)
        observed = body(actual[name], name)[0]
        if name == 'CCoupledSectionAnalysis':
            before = "' рабочий знак my соответствует +x·σz."
            assert expected.count(before) == 1
            expected = expected.replace(before, before.replace('σ', '?'))
            edits.append({'kind': 'known_inline_comment_encoding', 'only_comment': True})
        if name == 'CCoupledConcreteReturn':
            def TestConstantComment(code):
                """Учитывает только выравнивание известного комментария константы VBE."""
                lines = code.splitlines(keepends=True)
                for index, line in enumerate(lines):
                    if line.startswith('private const rounding_limit as double = '):
                        value, comment = line.split("'", 1)
                        lines[index] = value.rstrip() + " '" + comment
                return ''.join(lines)
            expected, observed = TestConstantComment(expected), TestConstantComment(observed)
        if name == 'modTestCoupledElasticState' and '0.25^(attempt-' in expected:
            assert expected.count('0.25^(attempt-') == 2
            expected = expected.replace('0.25^(attempt-', '0.25 ^ (attempt-')
            edits.append({'kind': 'known_exponent_operator_spacing',
                          'reason': 'caret before parenthesized exponent is an operator, not LongLong literal suffix'})
        expected, observed = normalize(expected)[0], normalize(observed)[0]
        def TestCommaSpaces(code):
            """Учитывает пробелы VBE после запятой, сохраняя строки, комментарии и все остальные символы."""
            return re.sub(r'"(?:[^"\n]|"")*"|\'[^\n]*|,[ \t]*',
                          lambda match: ', ' if match.group().startswith(',') else match.group(), code)
        if expected != observed and TestCommaSpaces(expected) == TestCommaSpaces(observed):
            edits.append({'kind': 'compiler_comma_spacing_only', 'strings_and_comments_preserved': True})
            expected, observed = TestCommaSpaces(expected), TestCommaSpaces(observed)
        if expected != observed:
            def TestLineTokens(code):
                """Сравнивает токены каждой строки; строки, комментарии, типы чисел и составные операторы остаются точными."""
                pattern = r'<number:[^>]+>|"(?:[^"\n]|"")*"|\'[^\n]*|[\w]+|<>|<=|>=|[^\s]'
                return [re.findall(pattern, line) for line in code.splitlines()]
            if TestLineTokens(expected) == TestLineTokens(observed):
                edits.append({'kind': 'compiler_intertoken_spacing_only',
                              'line_boundaries_strings_comments_typed_numbers_and_operators_preserved': True})
                expected = observed
        if expected != observed:
            print(name, '\n'.join(list(unified_diff(expected.splitlines(), observed.splitlines(), lineterm=''))[:50]))
            raise AssertionError(name)
        return edits
    for name in sorted(actual):
        if name not in compilation['Modules']:
            try:
                assert wrapped(actual[name], name)[0] == wrapped(original[name], name)[0], name
            except IndexError:
                assert actual[name] == original[name], name
            unchanged.append(name)
            continue
        paths = [file for directory in ('src', 'tests') for file in (root / directory).rglob(name + '.*')
                 if file.suffix in ('.cls', '.bas')]
        assert len(paths) == 1, name
        changes[name] = TestMatch(name, paths[0])
    coupled = []
    if (root / 'AllCoupledSources.json').exists():
        for relative, digest in json.loads((root / 'AllCoupledSources.json').read_bytes()).items():
            path = root / relative
            assert sha256(path.read_bytes()).hexdigest().upper() == digest
            TestMatch(path.stem, path)
            coupled.append(path.stem)
        assert len(coupled) == 33
    result = {'compiled_workbook_sha256': compilation['WorkbookSHA256'], 'compiled_project_storage': storage,
              'components': len(actual), 'imported': len(changes), 'unchanged': len(unchanged),
              'documented_compiler_conversions': changes, 'installed_workbook_changed': False,
              'checked_coupled_sources': sorted(coupled),
              'capacity_certified': False, 'full_acceptance_completed': False}
    output.write_text(json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
    print('ACTUAL VBA VERIFIED', result['components'], 'IMPORTED', result['imported'], 'UNCHANGED', result['unchanged'])


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, default=Path.cwd())
    parser.add_argument('--root', type=Path, required=True)
    parser.add_argument('--out', type=Path, required=True)
    args = parser.parse_args()
    TestAudit(args.repo.resolve(), args.root.resolve(), args.out.resolve())
