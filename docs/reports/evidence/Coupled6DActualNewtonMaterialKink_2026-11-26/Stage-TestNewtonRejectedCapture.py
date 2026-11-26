"""ДЛЯ ТЕСТОВ. Записывает реальные отклонённые поправки исходных восьми дуговых проб, не меняя вычисления."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent
base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold909ResidualArcPredictorBaseline_20261126'
root=base/'Cold911NewtonRejectedCapture_20261126'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'NewtonCaptureCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
def TestReplace(old,new):
    """Требует единственное точное совпадение исходной точки записи."""
    global code
    assert code.count(old)==1,old[:120]
    code=code.replace(old,new,1)
TestReplace('    SolveCorrection = mSystem.Solve(augmented, correction, size)', '''    If mUseSpatialHistory Then modTestCoupledElasticState.TestNewtonRejectedCapture 1, augmented, correction, field, mAccepted, mResidual, _
        Array(mAcceptedRevision, mIterations, mCurrentControlLambda, mGaugeStart, mReferenceModulus, mTargetControl, mArcLoadProjection, _
            Abs(CLng(mConsistentControl)), Abs(CLng(mRefinedControlTangent)))
    SolveCorrection = mSystem.Solve(augmented, correction, size)''')
anchor='                accepted = (candidateNorm <= candidateTolerance Or candidateNorm < norm * (1# - MIN_PROGRESS * fraction))'
TestReplace(anchor,anchor+'''
                If mControlActive And mUseSpatialHistory Then modTestCoupledElasticState.TestNewtonRejectedCapture 2, mMatrix, correction, _
                    mCandidate, mAccepted, mResidual, Array(fraction, candidateLambda, candidateNorm, candidateTolerance, ControlResidual(mCandidate), Abs(CLng(accepted)))''')
anchor='''        Loop
        If Not accepted Then
            If useSteelCorrection'''
TestReplace(anchor,'''        Loop
        If Not accepted Then
            If mControlActive And mUseSpatialHistory Then modTestCoupledElasticState.TestNewtonRejectedCapture 3, mMatrix, correction, _
                current, mAccepted, mResidual, Array(currentLambda, norm, tolerance, bestTrialNorm, Abs(CLng(hasValidTrial)), minimumFraction)
            If useSteelCorrection''')
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
code+='\n\n'+(here/'TestNewtonRejectedCapture.bas.txt').read_text('utf-8-sig')
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames @('CCoupledNewtonSolver','modTestCoupledElasticState') -EntryPointNames 'TestNewtonRejectedCapture'",1).replace('\n','\r\n').encode('utf-8-sig')
values.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():
    if name.endswith(('.cls','.bas')):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ActualRejectedNewtonCapture',actual_import_this_pass=['CCoupledNewtonSolver','modTestCoupledElasticState'],
 actual_component_count=256,no_additional_material_evaluations=True,linearization_before_linear_solve=True,
 only_already_valid_line_search_residuals_saved=True,no_numerical_or_material_changes=True,
 physical_capacity_certified=False,no_working_or_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'ACTUAL NEWTON CAPTURE;EIGHT ORIGINAL ARC CASES;NO NUMERICAL CHANGE')
