"""ДЛЯ ТЕСТОВ. Расширяет диагностическую ветвь до общего бюджета96 вызовов при неизменном рабочем алгоритме."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold823AutomaticLoadingWalk_20261125';root=base/'Cold825ExtendedLoadingWalk_20261125'
assert json.loads((parent/'AutomaticLoadingWalkProof.json').read_bytes())['new_accepted_equilibria']==8
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ExtendedLoadingCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='If newPoints >= 8 Or totalCalls + 3 > 96';assert code.count(anchor)==1
code=code.replace(anchor,'If newPoints >= 32 Or totalCalls + 3 > 96',1)
code=code.replace("' Проверяет до восьми новых равновесий общей рабочей операцией;", "' Проверяет до32 новых равновесий общей рабочей операцией;",1)
anchor='        sourceCount = mReverseFieldsTest.Count: revision = mReverseSolverTest.AcceptedRevision';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'\n        ReDim half(1 To 1): halfLambda = 0#',1)
anchor='        TestWriteAutomaticWalkPacket attempt, certificate, solved, projection\n';assert code.count(anchor)==1
code=code.replace(anchor,'',1)
anchor='            Set record = New Collection\n            record.Add sourceCount';assert code.count(anchor)==1
code=code.replace(anchor,'''            Dim sideSteel As CCoupledReinforcedResponse
            Set sideSteel = mReverseSidesTest(2)
            If sideSteel.TrialSteelUtilization > 1# Then Err.Raise 5, , "Численная ветвь превысила предел арматуры на отрицательной станции."
            Set sideSteel = mReverseSidesTest(4)
            If sideSteel.TrialSteelUtilization > 1# Then Err.Raise 5, , "Численная ветвь превысила предел арматуры на положительной станции."
            TestWriteAutomaticWalkPacket attempt, certificate, solved, projection, half, halfLambda
            Set record = New Collection
            record.Add sourceCount''',1)
anchor='        Else\n            mReverseSolverTest.DiscardStep\n        End If\n        Print #logHandle';assert code.count(anchor)==1
code=code.replace(anchor,'        Else\n            TestWriteAutomaticWalkPacket attempt, certificate, solved, projection, half, halfLambda\n            mReverseSolverTest.DiscardStep\n        End If\n        Print #logHandle',1)
anchor='Private Sub TestWriteAutomaticWalkPacket(ByVal attempt As Long, ByVal certificate As Variant, ByVal solved As Boolean, ByRef projection() As Double)';assert code.count(anchor)==1
code=code.replace(anchor,anchor[:-1]+', _\n        ByRef selectedHalf() As Double, ByVal selectedHalfLambda As Double)',1)
start=code.index('Private Sub TestWriteAutomaticWalkPacket(')
tail=code[start:];assert tail.count('magic = "NDMOPERARC000003"')==1
tail=tail.replace('magic = "NDMOPERARC000003"','magic = "NDMOPERARC000004"',1)
anchor='    TestWriteReverseBinaryVector handle, mReverseSolverTest.ReadPolishEffortTest\n    Close #handle';assert tail.count(anchor)==1
tail=tail.replace(anchor,'    TestWriteReverseBinaryVector handle, mReverseSolverTest.ReadPolishEffortTest\n    TestWriteReverseBinaryVector handle, selectedHalf\n    TestWriteReverseBinaryVector handle, Array(selectedHalfLambda)\n    Close #handle',1)
code=code[:start]+tail
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
values.update({'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():
    if n.endswith(('.cls','.bas')):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ExtendedWorkingAutomaticLoadingWalk',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
    all_working_methods_and_materials_exact823=True,maximum_new_points32=True,maximum_attempts32=True,maximum_calls96=True,
    attempts_per_revision4=True,all_three_station_trial_steel_limits_checked_before_accept=True,
    raw_selected_half_in_packet_v4=True,not_cold_capacity_search=True,no_future_field_or_hardcoded_station=True,
    physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root)
