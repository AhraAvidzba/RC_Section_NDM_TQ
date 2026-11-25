"""ДЛЯ ТЕСТОВ. Прослеживает фактические пути обновления бетонной точки15 вдоль подтверждённого свидетеля775."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold775ArcRootPath_20261125'
root = base / 'Cold776ArcPointRoute_20261125'
proof = json.loads((parent / 'ArcRootPathProof.json').read_bytes())
assert proof['zero_stress_and_residual_repeat_bitidentical'] and proof['source_states'] == 33
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0)
archive = root / 'PointRouteCompilationAncestor'; archive.mkdir()
for name in ('CompilePoint.log', 'CompilePointClosure.json'):
    (root / name).rename(archive / name)
prep = json.loads((root / 'Preparation.json').read_bytes())
name = 'src/Materials/CCoupledConcreteResponse.cls'
code = (root / name).read_text('utf-8-sig') + (here / 'TestPointUpdateRoute.cls.txt').read_text('utf-8-sig')
values = {name: code.replace('\n', '\r\n').encode('utf-8-sig')}
name = 'tests/modTestCoupledElasticState.bas'; code = (root / name).read_text('utf-8-sig')
code = code.replace('Option Explicit\n', 'Option Explicit\nPrivate mPointRouteHandleTest As Integer, mPointRouteQueryTest As Long\n', 1)
anchor = '    Dim index As Long, eta As Double, factor As Double, target As Double, kind As Long\n'
assert code.count(anchor) == 1
code = code.replace(anchor, '''    mPointRouteHandleTest = FreeFile
    Open mPolishRootTest & "\\Native\\ArcPoint15.Route.bin" For Binary Access Write As #mPointRouteHandleTest
    Dim routeMagic As String * 16: routeMagic = "NDMARCPOINT00001"
    Put #mPointRouteHandleTest, , routeMagic
'''+anchor, 1)
anchor = '        TestWriteArcRootNetTensors handle, field\n'
assert code.count(anchor) == 1
code = code.replace(anchor, '        mPointRouteQueryTest = index\n'+anchor, 1)
anchor = '    If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Читающая линия изменила принятую историю."'
assert code.count(anchor) == 1
code = code.replace(anchor, '    Close #mPointRouteHandleTest: mPointRouteHandleTest = 0\n'+anchor, 1)
anchor = '            For component = 1 To 6: position = position + 1: values(position) = stress(component): Next component\n'
assert code.count(anchor) == 1
extra = '''            If point = 15 Then
                Dim concrete As CCoupledConcreteResponse, trace As Collection, vector As Variant
                Select Case side
                    Case -1: Set concrete = mReverseSidesTest(1)
                    Case 0: Set concrete = mReverseConcreteTest
                    Case 1: Set concrete = mReverseSidesTest(3)
                End Select
                Set trace = concrete.ReadOperationalPointRouteTest(point, strain)
                TestWriteReverseBinaryVector mPointRouteHandleTest, Array(mPointRouteQueryTest, side, point)
                TestWriteReverseBinaryVector mPointRouteHandleTest, strain: TestWriteReverseBinaryVector mPointRouteHandleTest, stress
                For Each vector In trace: TestWriteReverseBinaryVector mPointRouteHandleTest, vector: Next vector
            End If
'''
code = code.replace(anchor, anchor + extra, 1)
assert max(map(len, code.splitlines())) < 1024
values[name] = code.replace('\n', '\r\n').encode('utf-8-sig')
compiler = (here / 'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1'] = compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames @('CCoupledConcreteResponse','modTestCoupledElasticState')", 1).replace('\n', '\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json'] = (parent / 'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json'] = (parent / 'CompilePointClosure.json').read_bytes()
values['TestPointUpdateRoute.cls.txt'] = (here / 'TestPointUpdateRoute.cls.txt').read_bytes()
values[Path(__file__).name] = Path(__file__).read_bytes()
for name, data in values.items():
    (root / name).write_bytes(data); prep['inputs'][name] = sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='ActualPoint15OperationalRoute', actual_import_this_pass=['CCoupledConcreteResponse','modTestCoupledElasticState'],
    all_working_material_methods_exact775=True, only_readonly_test_methods_added=True,
    bypass_cache_operational_answer_compared_with_normal_interface=True,
    operational_substeps_and_counters=True, independent_ordinary_return_on_same_history=True,
    history_not_accepted=True, physical_capacity_certified=False, no_working_promotion=True)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('776 POINT15 ON THREE STATIONS; ORDINARY/OPERATIONAL COUNTERS AND HISTORIES; WORKING METHODS EXACT775')
