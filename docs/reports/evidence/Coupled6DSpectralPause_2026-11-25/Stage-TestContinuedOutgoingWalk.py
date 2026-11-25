"""ДЛЯ ТЕСТОВ. Продолжает фактический69-точечный путь825 существующим алгоритмом, без повторения его поиска и без изменения материала."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold825ExtendedLoadingWalk_20261125';source=base/'Cold829OutgoingArcUncertainty_20261125'
root=base/'Cold833ContinuedOutgoingWalk_20261125'
proof=json.loads((source/'OutgoingArcUncertaintyProof.json').read_bytes())
assert proof['all_actual_source69_fields_exact825'] and proof['no_new_material_history_accepted']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0,source/'Input/Reverse1.bin')
archive=root/'ContinuedOutgoingCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='    TestRecentLoadingDirectionGuards\n    TestRunAutomaticLoadingWalk';assert code.count(anchor)==1
code=code.replace(anchor,'    TestRunContinuedOutgoingGuard\n    TestRunAutomaticLoadingWalk',1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет начало продолжения по принятой секущей без новой недавно достигнутой поверхности; будущая точка не передаётся.
Private Sub TestRunContinuedOutgoingGuard()
    Dim field() As Double, old() As Double, prediction() As Double, i As Long
    field = mReverseFieldsTest(mReverseFieldsTest.Count): old = mReverseFieldsTest(mReverseFieldsTest.Count - 1)
    prediction = field
    For i = 1 To mReverseSolverTest.PhysicalDegreeCount: prediction(i) = 2# * field(i) - old(i): Next i
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    Dim info As Variant
    If mReverseSolverTest.TryReadRecentLoadingDeparture(prediction, info) Then Err.Raise 5, , "Исходный69-точечный путь неожиданно заканчивается новой границей нагружения."
    If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Чтение исходной секущей изменило принятую историю."
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\RecentLoading.Guards.txt" For Output As #handle
    Print #handle, "NO_RECENT_BOUNDARY_EXPECTED;ARC_SECANT_USED;REVISION_UNCHANGED"
    Close #handle
End Sub
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
values.update({'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    'Input/Source825Attempts.json':(source/'Input/Source825Attempts.json').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():
    if n.endswith(('.cls','.bas')):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ContinueActual69StateOutgoingBranch',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
    numerical_kernel_and_materials_exact825=True,all69initial_fields_exact829_verified_from825_raw_packets=True,
    no_prescribed_future_field=True,maximum_new_points32=True,maximum_attempts32=True,maximum_calls96=True,
    all_three_station_steel_limits_before_accept=True,not_cold_capacity_search=True,
    physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root)
