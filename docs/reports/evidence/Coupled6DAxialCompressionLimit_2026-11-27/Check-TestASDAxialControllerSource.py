"""ДЛЯ ТЕСТОВ. Сверяет частную правку контроллера N или читающий наблюдатель с фактическим VBA и сохраняет остальные компоненты."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
from runpy import run_path
from argparse import ArgumentParser
import json, re

parser = ArgumentParser(description=__doc__)
parser.add_argument('--root', type=Path, required=True)
root = parser.parse_args().root
here = Path(__file__).resolve().parent
compiled = json.loads((root / 'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert compiled['Compiled'] and not compiled['ProductionChanged']
reuse = json.loads((root / 'CompiledSourceReuse.json').read_bytes())
expected_modules = ['CCoupledDeformationContinuation', 'modTestSelectableMaterialCore'] if reuse['variant'] in ('observe-peak', 'full-branch') else ['CCoupledDeformationContinuation']
assert compiled['Modules'] == expected_modules
assert json.loads((root / 'CompilePointClosure.json').read_text('utf-8-sig'))['State'] == 'Absent'
prep = json.loads((root / 'Preparation.json').read_bytes())
for name, digest in prep['inputs'].items():
    assert sha256((root / name).read_bytes()).hexdigest().upper() == digest, name
parent = Path(reuse['parent'])
before = (parent / 'docs/regression/Diagnostic.xlsm').read_bytes()
after = (root / 'docs/regression/Diagnostic.xlsm').read_bytes()
assert sha256(before).hexdigest().upper() == json.loads((root / 'ParentCodeUpdate.json').read_bytes())['WorkbookSHA256']
assert sha256(after).hexdigest().upper() == compiled['WorkbookSHA256']
helper, body, reader = run_path(str(here / 'Inspect-TestResolutionVBA.py'))['TestHelpers'](Path.cwd())
with ZipFile(BytesIO(before)) as handle:
    old_bin = handle.read('xl/vbaProject.bin')
with ZipFile(BytesIO(after)) as handle:
    new_bin = handle.read('xl/vbaProject.bin')
_, old = reader(helper['TestContainer'](old_bin))
_, new = reader(helper['TestContainer'](new_bin))
assert len(old) == 256 and set(old) == set(new)


def TestIdentifiers(text):
    """Допускает регистр имён VBE, сохраняя строки и комментарии при сравнении остальных компонентов."""
    return re.sub(r'"(?:[^"\n]|"")*"|\'[^\n]*|[A-Za-z_][\w]*',
                  lambda match: match.group() if match.group().startswith(('"', "'")) else match.group().lower(), text)


name = 'CCoupledDeformationContinuation'
for item in old:
    if item not in expected_modules:
        assert TestIdentifiers(old[item]) == TestIdentifiers(new[item]), item
normal = run_path(str(here / 'Canonical-TestCoreVBA.py'))['TestCanonical']
previous = TestIdentifiers(body(old[name], name)[0])
current = TestIdentifiers(body(new[name], name)[0])
if reuse['variant'] == 'budget':
    needle = 'private const probe_limit as long = 96'
    assert previous.count(needle) == 1
    previous = previous.replace(needle, 'private const probe_limit as long = 192', 1)
elif reuse['variant'] in ('peak', 'peak-current'):
    needle = "    If keepRight And UsesASDGeometricControl() Then Exit Function ' Локальный спад железобетона проходит как непроверенный; максимум не объявляется, критерий стали остаётся обязательным.\n"
    assert previous.count(needle.lower()) == 1
    previous = previous.replace(needle.lower(), '', 1)
elif reuse['variant'] == 'event':
    for point in ('point', 'high'):
        prefix = '            ' if point == 'high' else '        '
        needle = prefix + f'If Not Probe({point}, low.Lambda, keepYieldControl:=True, acceptedLambda:=low.Lambda) Then Exit Function'
        replacement = prefix + f'If Not Probe({point}, low.Lambda, keepYieldControl:=True, allowHybridRecovery:=mSolver.SupportsHybridRecovery, _\n'
        replacement += prefix + '        acceptedLambda:=low.Lambda) Then Exit Function'
        assert previous.count(needle.lower()) == 1
        previous = previous.replace(needle.lower(), replacement.lower(), 1)
elif reuse['variant'] == 'condition':
    needle = 'conditionHybridRecovery:=robustDeparture, referenceLambda:=recoveryReference, _'
    replacement = 'conditionHybridRecovery:=(robustDeparture Or (keepYieldControl And allowHybridRecovery)), referenceLambda:=recoveryReference, _'
    assert previous.count(needle.lower()) == 1
    previous = previous.replace(needle.lower(), replacement.lower(), 1)
elif reuse['variant'] == 'seed':
    previous = run_path(str(here / 'Patch-TestAxialSteelSeed.py'))['TestPatchSteelSeed'](previous, canonical=True)
elif reuse['variant'] == 'observe-peak':
    patch = run_path(str(here / 'Patch-TestPeakObservation.py'))
    previous = patch['TestPatchPeakObservation'](previous, canonical=True)
    staged = (root / 'tests/modTestSelectableMaterialCore.bas').read_text('utf-8-sig')
    staged_body = body(staged, 'modTestSelectableMaterialCore')[0]
    actual_body = body(new['modTestSelectableMaterialCore'], 'modTestSelectableMaterialCore')[0]
    assert normal(staged_body)[0] == normal(actual_body)[0]
    expected_module = patch['TestPatchPeakModule'](body(old['modTestSelectableMaterialCore'], 'modTestSelectableMaterialCore')[0])
    expected_module = body.__globals__['literal_safe_fold'](expected_module.rstrip())
    assert normal(expected_module)[0] == normal(actual_body)[0]
elif reuse['variant'] == 'observe-samples':
    previous = run_path(str(here / 'Patch-TestPeakObservation.py'))['TestPatchPeakSamples'](previous, canonical=True)
elif reuse['variant'] == 'adaptive-peak':
    previous = run_path(str(here / 'Patch-TestPeakPrecision.py'))['TestPatchPeakPrecision'](previous, canonical=True)
else:
    assert reuse['variant'] == 'full-branch'
    patch = run_path(str(here / 'Patch-TestFullCompressionBranch.py'))
    previous = patch['TestPatchFullBranchObservation'](previous, canonical=True)
    staged = (root / 'tests/modTestSelectableMaterialCore.bas').read_text('utf-8-sig')
    actual_body = body(new['modTestSelectableMaterialCore'], 'modTestSelectableMaterialCore')[0]
    assert normal(body(staged, 'modTestSelectableMaterialCore')[0])[0] == normal(actual_body)[0]
    expected_module = patch['TestPatchFullBranchModule'](body(old['modTestSelectableMaterialCore'], 'modTestSelectableMaterialCore')[0])
    expected_module = body.__globals__['literal_safe_fold'](expected_module.rstrip())
    assert normal(expected_module)[0] == normal(actual_body)[0]
assert normal(previous)[0] == normal(current)[0]
proof = dict(actual_components=256, all255_other_components_unchanged_except_identifier_case=True,
             changed_components_exact=True, controller_change_only=reuse['variant'],
             saved_book_sha256=compiled['WorkbookSHA256'], own_compiler_closed=True,
             full_equilibrium_acceptance_unchanged=True, physical_capacity_not_assumed=True)
if reuse['variant'] in ('observe-peak', 'full-branch'):
    proof.pop('all255_other_components_unchanged_except_identifier_case')
    proof['all254_other_components_unchanged_except_identifier_case'] = True
    proof['read_only_observation_components_exact'] = True
elif reuse['variant'] == 'observe-samples':
    proof['read_only_observation_components_exact'] = True
(root / 'EarlySourceProof.json').write_text(json.dumps(proof, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('PURE N CONTROLLER SOURCE VERIFIED', reuse['variant'], len(old) - len(expected_modules), 'OTHER COMPONENTS UNCHANGED')
