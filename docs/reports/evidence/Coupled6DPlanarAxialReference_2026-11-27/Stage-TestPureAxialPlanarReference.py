"""ДЛЯ ТЕСТОВ. Готовит чисто осевой контроль штатного плоского ядра без изменения численных классов и продолжения диаграмм."""
from pathlib import Path
from hashlib import sha256
import json, shutil, struct

repo = Path.cwd()
here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold1005PureBendingReference_20261126'
root = base / 'Cold1059PureAxialPlanarTension_20261127'
assert not root.exists()
checked = json.loads((parent / 'PureBendingReferenceProof.json').read_bytes())
assert checked['planar_reference_physical_limit_verified'] and checked['native']['own_process_closed']
compiled = json.loads((parent / 'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
book = (parent / 'docs/regression/Diagnostic.xlsm').read_bytes()
assert sha256(book).hexdigest().upper() == compiled['WorkbookSHA256']
for directory in ('src', 'tests', 'tools', 'Input', 'workbook'):
    shutil.copytree(parent / directory, root / directory)
for directory in ('Native', 'Temp', 'docs/regression'):
    (root / directory).mkdir(parents=True, exist_ok=True)
files = {
    'docs/regression/Diagnostic.xlsm': book,
    'ParentCodeUpdate.json': (parent / 'docs/regression/CodeUpdate.json').read_bytes(),
    'CompiledSourceReuse.json': (json.dumps(dict(parent=str(parent)), indent=2) + '\n').encode(),
    'Guard-IsolatedOffice.ps1': (parent / 'Guard-IsolatedOffice.ps1').read_bytes(),
    'Compile-TestExternalPoint.ps1': (here / 'Compile-TestExternalPoint.ps1').read_bytes(),
    'Run-TestPureAxialPlanarReference.ps1': (parent / 'Run-TestPureBendingReference.ps1').read_bytes(),
    Path(__file__).name: Path(__file__).read_bytes(),
    'Input/PlanarAxialBase.bin': struct.pack('<id', 1, 10000.),
}
method = (here / 'TestPureBendingPlanarReference.bas.txt').read_text('utf-8-sig')
for before, after in [
    ('TestRunPureBendingPlanarReference', 'TestRunPureAxialPlanarReference'),
    ('Находит штатный плоский предел Mx без продолжения диаграмм для геометрии последующего независимого контроля.',
     'Находит штатный плоский предел N по чисто осевой траектории без продолжения диаграмм.'),
    ('    cap.SolveByUltimateStrain section, concrete, steel, 0#, 1000000#, 0#',
     '    cap.SolveByUltimateLoadPath section, concrete, steel, 0#, nBase, 0#, 0#, 0#, 0#'),
    ('PlanarReference.txt', 'PlanarAxialReference.txt'),
    ('PlanarReference.bin', 'PlanarAxialReference.bin'),
    ('NDMPLANEMOM00001', 'NDMPLANEN0000001'),
    ('"PLANAR_REFERENCE: Mx_Nmm=" & CStr(cap.MxUltimate)', '"PLANAR_AXIAL_REFERENCE: N_N=" & CStr(cap.NUltimate)'),
    ('TOTAL_PURE_BENDING_REFERENCE', 'TOTAL_PURE_AXIAL_PLANAR_REFERENCE'),
]:
    assert before in method, before
    method = method.replace(before, after)
anchor = '    Dim cap As New CCapacitySolver\n'
assert method.count(anchor) == 1
method = method.replace(anchor, '''    Dim inputHandle As Integer: inputHandle = FreeFile
    Open folder & "\\Input\\PlanarAxialBase.bin" For Binary Access Read As #inputHandle
    Dim inputCount As Long, nBase As Double
    If LOF(inputHandle) <> 12 Then Err.Raise 5, , "Некорректная длина осевой нагрузки теста."
    Get #inputHandle, , inputCount: Get #inputHandle, , nBase
    Close #inputHandle: inputHandle = 0
    If inputCount <> 1 Or nBase = 0# Then Err.Raise 5, , "Тесту нужна одна ненулевая осевая нагрузка."
''' + anchor, 1)
module_path = root / 'tests/modTestCoupledElasticState.bas'
module = module_path.read_text('utf-8-sig')
anchor = 'Public Function TestRunPureBendingPlanarReference('
assert module.count(anchor) == 1 and 'Public Function TestRunPureAxialPlanarReference(' not in module
module = module.replace(anchor, method + '\n\n' + anchor, 1)
assert max(map(len, module.splitlines())) < 1024
files['tests/modTestCoupledElasticState.bas'] = module.replace('\n', '\r\n').encode('utf-8-sig')
files['TestPureAxialPlanarReference.bas.txt'] = method.replace('\n', '\r\n').encode('utf-8-sig')
files['Run-TestPureAxialPlanarReference.ps1'] = files['Run-TestPureAxialPlanarReference.ps1'].decode('utf-8-sig').replace(
    'TestRunPureBendingPlanarReference', 'TestRunPureAxialPlanarReference').encode('utf-8-sig')
files['Input/PlanarAxialControl.json'] = (json.dumps(dict(
    geometry='Same 180x120mm rectangle and two longitudinal bars as checked planar bending reference1005',
    base_N_N=10000., positive_N_means='tension', load_path='lambda*N',
    reference_method='CCapacitySolver.SolveByUltimateLoadPath;stock pure axial redirect to LoadMultiplier',
    concrete_diagram='ThreeLine', concrete_tension='Ignore', steel_diagram='TwoLine',
    diagram_extension=False, no_coupled_kernel_used=True,
    spatial_materials_differ=True, physical_capacity_not_assumed=True), sort_keys=True, indent=2) + '\n').encode()
for name, data in files.items():
    (root / name).write_bytes(data)
prep = dict(parent=str(parent), directed_scope=dict(case='PureAxialPlanarTension',
            actual_import_this_pass=['modTestCoupledElasticState'], all_numerical_classes_unchanged=True,
            only_new_test_function=True, no_artificial_diagram_extension=True),
            inputs={p.relative_to(root).as_posix(): sha256(p.read_bytes()).hexdigest().upper()
                    for p in root.rglob('*') if p.is_file() and p.relative_to(root).as_posix() != 'docs/regression/Diagnostic.xlsm'})
(root / 'Preparation.json').write_text(json.dumps(prep, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('PLANAR PURE N TENSION PREPARED', root, 'STOCK NUMERICAL CLASSES;NO DIAGRAM EXTENSION')
