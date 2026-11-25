"""ДЛЯ ТЕСТОВ. Читает новые границы CDP до принятия каждого точного исходного поля; решатель и история не меняются."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from datetime import datetime
import json

here=Path(__file__).resolve().parent
base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold815ActiveLoadingNeighborhood_20261125'
root=base/'Cold817RecentLoadingBoundary_20261125'
assert json.loads((parent/'ActiveLoadingCoordinateProof.json').read_bytes())['pair']['verified']
if not root.exists():
    run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
    archive=root/'RecentBoundaryCompilationAncestor';archive.mkdir()
    for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
else:
    assert (root/'RecentBoundaryCompilationAncestor').is_dir() and not (root/'CompilePoint.log').exists()
    assert (root/'tests/modTestCoupledElasticState.bas').read_bytes()==(parent/'tests/modTestCoupledElasticState.bas').read_bytes()
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
assert code.count('    TestRunActiveLoadingCoordinate\n')==1
code=code.replace('    TestRunActiveLoadingCoordinate\n','',1)
anchor='        mReverseSolverTest.AcceptStep\n        mReverseOperatorsTest.LoadEquations'
assert code.count(anchor)==2
code=code.replace(anchor,'''        TestReadRecentLoadingBoundary index, field
        mReverseSolverTest.AcceptStep
        mReverseOperatorsTest.LoadEquations''')
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает старое и принимаемое поле на одной принятой пластической истории до её изменения.
Private Sub TestReadRecentLoadingBoundary(ByVal index As Long, ByRef field() As Double)
    Dim old() As Double: mReverseSolverTest.CopyAccepted old
    Dim point As Long, side As Long, left As Double, right As Double, available As Boolean
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\RecentLoading.Boundaries.csv" For Append As #handle
    For point = 1 To mReverseOperatorsTest.PointCount
        For side = -1 To 1
            available = mReverseSolverTest.TryReadLoadingSurface(point, side, old, left)
            available = mReverseSolverTest.TryReadLoadingSurface(point, side, field, right) And available
            If available Then Print #handle, CStr(index) & ";" & CStr(point) & ";" & CStr(side) & ";" & _
                TestOnsetNumber(left) & ";" & TestOnsetNumber(right)
        Next side
    Next point
    Close #handle
End Sub
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=(here/'Compile-TestExternalPoint.ps1').read_bytes()
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for n,data in values.items():
    (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ReadRecentLoadingBoundaryBeforeCommit',actual_import_this_pass=['modTestCoupledElasticState'],
    actual_component_count=256,all_working_methods_and_source43_fields_exact815=True,
    no_new_NDS_or_capacity_search=True,no_material_acceptance_beyond_source_replay=True,
    physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=Path('docs/reports/Coupled6DResumeState_2026-11-24.json');state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,
    pause_request_revoked=True,pause_checkpoint_ready=False,goal_pause_tool_pending=False,
    active_jobs=[],all_started_stages_terminal=True,
    next_action='817: verify recent loading-boundary detection in actual source export; 816 instrumented an unused export, no boundary evidence from816; full goal unchanged')
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (path,here/'Continuation666_Resume20261123.json'):target.write_bytes(raw)
print(root)
