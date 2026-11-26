"""ДЛЯ ТЕСТОВ. Проверяет по одному изменению контроллера N: бюджет, уточнение максимума либо численный резерв события стали."""
from pathlib import Path
from hashlib import sha256
from argparse import ArgumentParser
import json, shutil, struct

parser = ArgumentParser(description=__doc__)
parser.add_argument('--variant', choices=('budget', 'peak', 'event', 'condition', 'seed', 'peak-current', 'observe-peak', 'observe-samples', 'adaptive-peak', 'full-branch'), required=True)
variant = parser.parse_args().variant
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent_name, root_name = {
    'budget': ('Cold1041ASDPureNHybrid_20261127', 'Cold1042ASDPureNBudget_20261127'),
    'peak': ('Cold1040ASDPureNCompression_20261127', 'Cold1043ASDPureNPeak_20261127'),
    'event': ('Cold1045ASDPureNScale10kN_20261127', 'Cold1046ASDPureNEventRecovery_20261127'),
    'condition': ('Cold1046ASDPureNEventRecovery_20261127', 'Cold1047ASDPureNEventCondition_20261127'),
    'seed': ('Cold1047ASDPureNEventCondition_20261127', 'Cold1048ASDPureNEventSeed_20261127'),
    'peak-current': ('Cold1048ASDPureNEventSeed_20261127', 'Cold1051ASDPureNCurrentCompression_20261127'),
    'observe-peak': ('Cold1051ASDPureNCurrentCompression_20261127', 'Cold1052ASDPureNPeakObservation_20261127'),
    'observe-samples': ('Cold1052ASDPureNPeakObservation_20261127', 'Cold1053ASDPureNPeakSamples_20261127'),
    'adaptive-peak': ('Cold1053ASDPureNPeakSamples_20261127', 'Cold1055ASDPureNAdaptivePeak_20261127'),
    'full-branch': ('Cold1055ASDPureNAdaptivePeak_20261127', 'Cold1056ASDPureNFullCompressionBranch_20261127')
}[variant]
parent, root = base / parent_name, base / root_name
assert not root.exists()
compiled = json.loads((parent / 'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
book = (parent / 'docs/regression/Diagnostic.xlsm').read_bytes()
assert sha256(book).hexdigest().upper() == compiled['WorkbookSHA256']
assert json.loads((parent / 'MainCapacityExecutionProof.json').read_bytes())['native']['own_process_closed']
root.mkdir()
for directory in ('src', 'tests', 'tools', 'workbook', 'Input'):
    shutil.copytree(parent / directory, root / directory)
for directory in ('Native', 'Temp', 'docs/regression'):
    (root / directory).mkdir(parents=True, exist_ok=True)
for name in ('Guard-IsolatedOffice.ps1', 'Run-TestASDSpatialMainCapacity.ps1'):
    shutil.copyfile(parent / name, root / name)
(root / 'docs/regression/Diagnostic.xlsm').write_bytes(book)
shutil.copyfile(parent / 'docs/regression/CodeUpdate.json', root / 'ParentCodeUpdate.json')
name = 'src/Solver/CCoupledDeformationContinuation.cls'
code = (root / name).read_text('utf-8-sig')
if variant == 'budget':
    needle = 'Private Const PROBE_LIMIT As Long = 96'
    assert code.count(needle) == 1
    code = code.replace(needle, 'Private Const PROBE_LIMIT As Long = 192', 1)
elif variant in ('peak', 'peak-current'):
    needle = "    If keepRight And UsesASDGeometricControl() Then Exit Function ' Локальный спад железобетона проходит как непроверенный; максимум не объявляется, критерий стали остаётся обязательным.\n"
    assert code.count(needle) == 1
    code = code.replace(needle, '', 1)
elif variant == 'event':
    for point in ('point', 'high'):
        needle = f'        '+('    ' if point == 'high' else '')+f'If Not Probe({point}, low.Lambda, keepYieldControl:=True, acceptedLambda:=low.Lambda) Then Exit Function'
        assert code.count(needle) == 1
        prefix = '            ' if point == 'high' else '        '
        replacement = prefix + f'If Not Probe({point}, low.Lambda, keepYieldControl:=True, allowHybridRecovery:=mSolver.SupportsHybridRecovery, _\n'
        replacement += prefix + '        acceptedLambda:=low.Lambda) Then Exit Function'
        code = code.replace(needle, replacement, 1)
elif variant == 'condition':
    needle = 'conditionHybridRecovery:=robustDeparture, referenceLambda:=recoveryReference, _'
    replacement = 'conditionHybridRecovery:=(robustDeparture Or (keepYieldControl And allowHybridRecovery)), referenceLambda:=recoveryReference, _'
    assert code.count(needle) == 1
    code = code.replace(needle, replacement, 1)
elif variant == 'seed':
    from runpy import run_path
    code = run_path(str(Path(__file__).resolve().parent / 'Patch-TestAxialSteelSeed.py'))['TestPatchSteelSeed'](code)
elif variant == 'observe-peak':
    from runpy import run_path
    patch = run_path(str(Path(__file__).resolve().parent / 'Patch-TestPeakObservation.py'))
    code = patch['TestPatchPeakObservation'](code)
    module = root / 'tests/modTestSelectableMaterialCore.bas'
    updated = patch['TestPatchPeakModule'](module.read_text('utf-8-sig'))
    module.write_bytes(updated.replace('\n', '\r\n').encode('utf-8-sig'))
elif variant == 'observe-samples':
    from runpy import run_path
    code = run_path(str(Path(__file__).resolve().parent / 'Patch-TestPeakObservation.py'))['TestPatchPeakSamples'](code)
elif variant == 'adaptive-peak':
    from runpy import run_path
    code = run_path(str(Path(__file__).resolve().parent / 'Patch-TestPeakPrecision.py'))['TestPatchPeakPrecision'](code)
else:
    from runpy import run_path
    patch = run_path(str(Path(__file__).resolve().parent / 'Patch-TestFullCompressionBranch.py'))
    code = patch['TestPatchFullBranchObservation'](code)
    module = root / 'tests/modTestSelectableMaterialCore.bas'
    updated = patch['TestPatchFullBranchModule'](module.read_text('utf-8-sig'))
    module.write_bytes(updated.replace('\n', '\r\n').encode('utf-8-sig'))
(root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
if variant == 'peak-current':
    (root / 'Input/CapacityLoad.bin').write_bytes(struct.pack('<i6di', 6, -100000., 0., 0., 0., 0., 0., 1))
compiler = (base / 'Cold1041ASDPureNHybrid_20261127/Compile-TestASDHybridRecovery.ps1').read_text('utf-8-sig')
assert compiler.count("-ModuleNames 'CCoupledNewtonSolver'") == 1
compiler = compiler.replace("-ModuleNames 'CCoupledNewtonSolver'", "-ModuleNames 'CCoupledDeformationContinuation'", 1)
if variant in ('observe-peak', 'full-branch'):
    compiler = compiler.replace("-ModuleNames 'CCoupledDeformationContinuation'", "-ModuleNames @('CCoupledDeformationContinuation','modTestSelectableMaterialCore')", 1)
(root / 'Compile-TestASDAxialController.ps1').write_bytes(compiler.replace('\n', '\r\n').encode('utf-8-sig'))
(root / 'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent), variant=variant), indent=2) + '\n', encoding='utf-8')
shutil.copyfile(Path(__file__), root / Path(__file__).name)
prep = dict(directed_scope=dict(case='ASD_pure_N_controller_' + variant, changed_component='CCoupledDeformationContinuation',
            change={'budget': 'probe budget96to192 only',
                    'peak': 'remove private early-return before existing peak refinement only',
                    'event': 'allow existing same-equation QR fallback in two steel-event refinement probes only',
                    'condition': 'use existing diagonal scaling and fresh Jacobian for steel-event QR fallback; same control and no extra load-advance requirement',
                    'seed': 'use actual solved low/high fields only as initial guesses for steel-event probes; re-solve from current committed history',
                    'peak-current': 'restore existing peak refinement for pure compression on verified192-probe seed controller',
                    'observe-peak': 'read-only scalar observer of peak refinement; numerical expressions unchanged',
                    'observe-samples': 'read-only capture of actual shoulder samples before peak recognition',
                    'adaptive-peak': 'tighten peak-equilibrium probes from adjoint uncertainty; unchanged lambda tolerance and physical model',
                    'full-branch': 'read-only export of complete accepted post-peak branch before final maximum verification'}[variant],
            all_materials_and_equilibrium_tolerances_unchanged=True, physical_capacity_not_assumed=True),
            inputs={p.relative_to(root).as_posix(): sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
                    if p.is_file() and p.relative_to(root).as_posix() != 'docs/regression/Diagnostic.xlsm'})
(root / 'Preparation.json').write_text(json.dumps(prep, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('PURE N CONTROLLER PREPARED', root, variant, 'MATERIAL AND EQUILIBRIUM CRITERIA UNCHANGED')
