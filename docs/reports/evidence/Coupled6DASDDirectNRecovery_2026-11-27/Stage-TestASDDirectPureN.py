"""ДЛЯ ТЕСТОВ. Готовит шесть прямых холодных осевых расчётов ASD с читающим экспортом истории и неизменным численным алгоритмом."""
from pathlib import Path
from hashlib import sha256
import json, struct, shutil

repo = Path.cwd()
here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold1056ASDPureNFullCompressionBranch_20261127'
root = base / 'Cold1061ASDPureNDirectBaseline_20261127'
assert not root.exists()
proof = json.loads((parent / 'AxialCompressionLimitProof.json').read_bytes())
assert proof['peak_load_capacity_confirmed_on_followed_branch']
compiled = json.loads((parent / 'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
book = (parent / 'docs/regression/Diagnostic.xlsm').read_bytes()
assert sha256(book).hexdigest().upper() == compiled['WorkbookSHA256']
for directory in ('src', 'tests', 'tools', 'Input', 'workbook'):
    shutil.copytree(parent / directory, root / directory)
for directory in ('Native', 'Temp', 'docs/regression'):
    (root / directory).mkdir(parents=True, exist_ok=True)
module = (parent / 'tests/modTestSelectableMaterialCore.bas').read_text('utf-8-sig')
start = module.index('Public Function TestRunASDSpatialMainCapacity(')
end = module.index('\nEnd Function', start)
function = module[start:end]
setup = function[function.index('    Dim geometry As CGeometryRegion'):function.index('    loads.Initialize')]
assert setup.count('concreteLaw:="ASD"') == 1
start = module.index('Public Sub TestCaptureMainCapacity(')
end = module.index('\nEnd Sub', start)
capture = module[start:end]
body = capture[capture.index('    Dim history As Collection'):]
needle = '    Set history = controlled.AcceptedFields: Set trace = controlled.SavedTrace'
assert body.count(needle) == 1
body = body.replace(needle, '    Set history = continuation.AcceptedFields: Set trace = continuation.SavedTrace', 1)
method = (here / 'TestASDDirectPureN.bas.txt').read_text('utf-8-sig')
assert method.count("    'TEST_DIRECT_SETUP_INSERT'") == method.count("    'TEST_DIRECT_HISTORY_INSERT'") == 1
method = method.replace("    'TEST_DIRECT_SETUP_INSERT'", setup.rstrip()).replace("    'TEST_DIRECT_HISTORY_INSERT'", body)
assert 'Public Function TestRunASDDirectPureN(' not in module
module = module.rstrip() + '\n\n' + method
analysis_path = 'src/Solver/CCoupledSectionAnalysis.cls'
analysis = (parent / analysis_path).read_text('utf-8-sig')
anchor = '    Set acceptedTrace = continuation.SavedTrace\n'
assert analysis.count(anchor) == 1
analysis = analysis.replace(anchor, '    modTestSelectableMaterialCore.TestCaptureDirectHistory continuation, mOperators, path\n' + anchor, 1)
compiler = (parent / 'Compile-TestASDAxialController.ps1').read_text('utf-8-sig')
before = "@('CCoupledDeformationContinuation','modTestSelectableMaterialCore')"
assert compiler.count(before) == 1
compiler = compiler.replace(before, "@('CCoupledSectionAnalysis','modTestSelectableMaterialCore')", 1)
runner = (parent / 'Run-TestASDSpatialMainCapacity.ps1').read_text('utf-8-sig')
assert runner.count('TestRunASDSpatialMainCapacity') == 1
runner = runner.replace('TestRunASDSpatialMainCapacity', 'TestRunASDDirectPureN', 1)
loads = [50000., 90000., 98000., 100000., -700000., -750000.]
files = {
    'docs/regression/Diagnostic.xlsm': book,
    'ParentCodeUpdate.json': (parent / 'docs/regression/CodeUpdate.json').read_bytes(),
    'CompiledSourceReuse.json': (json.dumps(dict(parent=str(parent), variant='read-only-direct'), indent=2) + '\n').encode(),
    'Guard-IsolatedOffice.ps1': (parent / 'Guard-IsolatedOffice.ps1').read_bytes(),
    'Compile-TestASDDirectPureN.ps1': compiler.replace('\n', '\r\n').encode('utf-8-sig'),
    'Run-TestASDDirectPureN.ps1': runner.replace('\n', '\r\n').encode('utf-8-sig'),
    'tests/modTestSelectableMaterialCore.bas': module.replace('\n', '\r\n').encode('utf-8-sig'),
    analysis_path: analysis.replace('\n', '\r\n').encode('utf-8-sig'),
    'TestASDDirectPureN.generated.bas.txt': method.replace('\n', '\r\n').encode('utf-8-sig'),
    'TestASDDirectPureN.bas.txt': (here / 'TestASDDirectPureN.bas.txt').read_bytes(),
    Path(__file__).name: Path(__file__).read_bytes(),
    'Input/DirectLoads.bin': struct.pack('<i6d', 6, *loads),
    'Input/DirectControl.json': (json.dumps(dict(loads_N=loads, independent_zero_history_each_case=True,
        same_section_materials_as_1056=True, only_ASD=True, CDP_deferred_by_user=True,
        load_path='lambda*N', tolerances_unchanged=True, physical_acceptance_not_assumed=True), indent=2) + '\n').encode(),
}
for index, force in enumerate(loads, 1):
    case = root / f'Case{index:02d}'
    for directory in ('Input', 'Native', 'docs/regression'):
        (case / directory).mkdir(parents=True, exist_ok=True)
    files[f'Case{index:02d}/Input/CapacityLoad.bin'] = struct.pack('<i6di', 6, force, 0., 0., 0., 0., 0., 1)
    files[f'Case{index:02d}/Input/MaterialContract.json'] = (parent / 'Input/MaterialContract.json').read_bytes()
for name, data in files.items():
    (root / name).write_bytes(data)
assert max(map(len, module.splitlines())) < 1024
prep = dict(parent=str(parent), directed_scope=dict(case='ASDDirectPureNBaseline', actual_import_this_pass=[
            'CCoupledSectionAnalysis', 'modTestSelectableMaterialCore'], only_reading_observer_and_test_added=True,
            all_numerical_algorithms_unchanged=True, cases=loads, no_CDP_run=True),
            inputs={p.relative_to(root).as_posix(): sha256(p.read_bytes()).hexdigest().upper()
                    for p in root.rglob('*') if p.is_file() and p.relative_to(root).as_posix() != 'docs/regression/Diagnostic.xlsm'})
(root / 'Preparation.json').write_text(json.dumps(prep, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('ASD DIRECT PURE N PREPARED', root, loads, 'NUMERICAL ALGORITHM UNCHANGED;NO CDP')
