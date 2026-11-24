"""ДЛЯ ТЕСТОВ. Ужесточает только локальные критерии бетона после точного принятого префикса; глобальные допуски не меняет."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold729TighterLocalShoulder_20261124';parent=base/'Cold728FastShoulderWitness_20261124'
proof=json.loads((parent/'ShoulderWitnessProof.json').read_bytes())
assert proof['native']['own_process_closed'] and proof['direct_solved']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'TighterLocalCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
values={};prep=json.loads((root/'Preparation.json').read_bytes())
name='src/Materials/CCoupledConcreteResponse.cls';code=(root/name).read_text('utf-8-sig')
lines=[]
for line in code.splitlines():
    if not line.startswith('Private Const '):
        line=re.sub(r'\bSTATE_TOLERANCE\b','mStateToleranceTest',line)
        line=re.sub(r'\bYIELD_TOLERANCE_RATIO\b','mYieldToleranceTest',line)
    lines.append(line)
code='\n'.join(lines)+'\n'
anchor='Option Explicit\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'''Private mStateToleranceTest As Double, mYieldToleranceTest As Double
Private mPrecisionFactorTest As Double, mPrecisionEpochTest As Long
''',1)
anchor='    Set mASDFlow = Nothing\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'    mStateToleranceTest = STATE_TOLERANCE: mYieldToleranceTest = YIELD_TOLERANCE_RATIO\n',1)
anchor='    Dim i As Long, moved As Boolean\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'''    mStateToleranceTest = STATE_TOLERANCE: mYieldToleranceTest = YIELD_TOLERANCE_RATIO
    If mPrecisionFactorTest > 0# And mTrialEpoch >= mPrecisionEpochTest Then
        mStateToleranceTest = STATE_TOLERANCE * mPrecisionFactorTest
        mYieldToleranceTest = YIELD_TOLERANCE_RATIO * mPrecisionFactorTest
    End If
''',1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ужесточает локальную приёмку после выбранного числа принятых полей; исходная история и критерии сечения не меняются.
Friend Sub SetLocalPrecisionTest(ByVal factor As Double, ByVal startEpoch As Long)
    If factor <= 0# Or factor > 1# Or startEpoch < 1 Or mPointCount < 1 Then Err.Raise 5, , "Неверная точность локального теста."
    If Not mASDFlow Is Nothing Then Err.Raise 5, , "Тест локальной точности относится к CDP."
    mPrecisionFactorTest = factor: mPrecisionEpochTest = startEpoch
    Dim point As Long
    For point = 1 To mPointCount: mCachedEpoch(point) = 0: Next point
End Sub
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='        Set mReplayMaterialsTest = New Collection\n';assert code.count(anchor)==1
code=code.replace(anchor,'''        For index = LBound(concrete) To UBound(concrete)
            concrete(index).SetLocalPrecisionTest 0.01, fields.Count + 1
        Next index
'''+anchor,1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames @('CCoupledConcreteResponse','modTestCoupledElasticState')")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='TighterLocalPrecisionForSameShoulder',actual_import_this_pass=['CCoupledConcreteResponse','modTestCoupledElasticState'],
    local_tolerance_factor=.01,local_tightening_only_after_original_prefix=True,original_accepted_material_history_preserved=True,
    trial_cache_invalidated_on_precision_change=True,global_quality_criteria_and_30_iterations_unchanged=True,
    original_uniform_adjoint_estimator_unchanged=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('729 LOCAL MATERIAL TOLERANCES100X TIGHTER ONLY AFTER EXACT52 PREFIX; GLOBAL CRITERIA UNCHANGED')
