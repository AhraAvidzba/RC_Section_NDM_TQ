"""ДЛЯ ТЕСТОВ. Использует совместно матрицу и остаток действительной точки новой ветви; приёмка сравнивается с исходной нормой."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold915ProspectiveBranchRecovery_20261126';root=base/'Cold918AffineBranchRecovery_20261126'
proof=json.loads((base/'Cold916ProspectiveAffineRoots_20261126/ProspectiveAffineRootsProof.json').read_bytes())
assert proof['independently_equilibrated_affine_predictions']==5
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'AffineRecoveryCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
def TestReplace(old,new):
    """Требует точный исходный блок перед изменением одной согласованной линейной пробы."""
    global code
    assert code.count(old)==1,old[:100];code=code.replace(old,new,1)
TestReplace("' Остаток, строка управления и нагрузка относятся к исходной точке; новое поле проходит обычный поиск доли.",
            "' Матрица и остаток относятся к одной пробной точке новой ветви; улучшение проверяется относительно исходной нормы.")
TestReplace('    Dim prospective() As Double, correction() As Double, trial() As Double, fraction As Double',
            '    Dim correction() As Double, trial() As Double, fraction As Double')
TestReplace('''    prospective = mMatrix
    SetControlledLoad currentLambda, load
    If Not Evaluate(current, False, recordTrial:=False) Then GoTo RestoreFailure
    mMatrix = prospective ' Меняется только линеаризация; остаток и управляющая строка принадлежат текущей точке.
    If Not SolveCorrection(current, correction) Then GoTo RestoreFailure''',
'''    If Not SolveCorrection(oldCandidate, correction) Then GoTo RestoreFailure ' Нагрузка и управляющий остаток принадлежат этой же пробной точке.''')
TestReplace('''        trial = current
        For i = 1 To mDegreeCount: trial(i) = current(i) + fraction * correction(i): Next i
        trialLambda = currentLambda + fraction * correction(mDegreeCount + 1)''',
'''        trial = oldCandidate
        For i = 1 To mDegreeCount: trial(i) = oldCandidate(i) + fraction * correction(i): Next i
        trialLambda = oldLambda + fraction * correction(mDegreeCount + 1)''')
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledNewtonSolver'",1).replace('\n','\r\n').encode('utf-8-sig')
values.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='AffineProspectiveBranchNewtonRecovery',actual_import_this_pass=['CCoupledNewtonSolver'],
 actual_component_count=256,material_and_force_acceptance_unchanged=True,scalar_control_and_load_path_unchanged=True,
 matrix_and_rhs_from_same_postyield_trial=True,end_acceptance_compares_original_norm=True,
 every_recovery_attempt_consumes_two_primary_iterations=True,primary30_global96_unchanged=True,
 no_material_history_acceptance_in_recovery=True,failed_recovery_restores_original_context=True,
 no_working_or_release_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'MATCHED POSTYIELD MATRIX/RHS;ACCEPTANCE AGAINST ORIGINAL NORM;SAME BUDGET')
