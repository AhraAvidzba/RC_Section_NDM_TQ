"""ДЛЯ ТЕСТОВ. Проверяет полный материальный корень от принятой ветви до обычной итерации дилатансии."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold711PriorDilationCurve_20261124';root=base/'Cold722PriorBranchClosure_20261124'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'PriorBranchCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
(root/'Matrix/Native').mkdir(parents=True,exist_ok=True)
values={};prep=json.loads((root/'Preparation.json').read_bytes())
path='src/Materials/CCoupledConcreteResponse.cls';code=(root/path).read_text('utf-8-sig')
anchor='    If UpdatePoint(strain, committed, updated, stress, virginElastic) Then\n';assert code.count(anchor)==1
code=code.replace(anchor,'''    If CanPredictDilationBranch(strain, committed) Then
        mClosureRequestsTest = mClosureRequestsTest + 1
        If TryDilationClosure(strain, committed, updated, stress, True) Then
            mClosureSuccessesTest = mClosureSuccessesTest + 1
            StoreStatePosition updated, strain, stress
            virginElastic = False: updated.ElasticTangentExact = False: substeps = 1
            UpdateWithSubsteps = True
            Exit Function
        End If
    End If
'''+anchor,1)
anchor='ByRef result As TConcreteHistory, ByRef stress() As Double) As Boolean\n    On Error GoTo ArithmeticFailure'
assert code.count(anchor)==1
code=code.replace(anchor,'ByRef result As TConcreteHistory, ByRef stress() As Double, Optional ByVal preferCommitted As Boolean = False) As Boolean\n    On Error GoTo ArithmeticFailure',1)
anchor='    divisions = 2\n    Do\n        prior = committed: available = False\n';assert code.count(anchor)==1
code=code.replace(anchor,'''    If preferCommitted Then
        guide = committed
        For component = 1 To N_COMPONENTS: guideStress(component) = committed.Stress(component): Next component
        available = True
    Else
'''+anchor,1)
anchor='    Loop While divisions <= MAX_LOCAL_DIVISIONS\n    If Not available Then Exit Function\n';assert code.count(anchor)==1
code=code.replace(anchor,'    Loop While divisions <= MAX_LOCAL_DIVISIONS\n    End If\n    If Not available Then Exit Function\n',1)
code+='''
' Продолжает растянутую негидростатическую пластическую ветвь по принятой истории; новый полный шаг заново проверяет поверхность и закон потока.
Private Function CanPredictDilationBranch(ByRef strain() As Double, ByRef committed As TConcreteHistory) As Boolean
    If Not mASDFlow Is Nothing Or committed.Tension <= 0# Then Exit Function
    If strain(1) + strain(2) + strain(3) <= committed.Strain(1) + committed.Strain(2) + committed.Strain(3) Then Exit Function
    Dim trace As Double, q As Double, high As Double, middle As Double, low As Double, damage As Double
    CoupledStressInvariants committed.Stress, trace, q, high, middle, low
    If trace <= 0# Then Exit Function
    damage = (1# - PositiveStressFraction(high, middle, low) * committed.TensionDamage) * (1# - committed.CompressionDamage)
    If damage <= 0# Then Exit Function
    CanPredictDilationBranch = (q > TENSILE_AXIS_BAND_RATIO * mFc * damage + PLASTIC_ROUNDOFF * GeomMax(Abs(high), Abs(low)))
End Function
'''
assert max(map(len,code.splitlines()))<1024
values[path]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledConcreteResponse'")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='PriorBranchFullStepClosure',actual_import_this_pass=['CCoupledConcreteResponse'],
    original21_targets_and_test_expectations_exact=True,full_step_yield_and_flow_acceptance_unchanged=True,
    candidate_full_step_initialized_from_committed_history=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('722 PRIOR BRANCH FULL STEP; ORIGINAL21 TARGETS AND TEST MODULE; SAME MATERIAL EQUATIONS')
