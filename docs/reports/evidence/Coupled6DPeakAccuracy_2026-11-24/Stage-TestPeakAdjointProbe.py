"""ДЛЯ ТЕСТОВ. Записывает точные сопряжённые матрицы и исходную причину отказа повторения плеч без изменения численных решений713."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold715PeakAdjointProbe_20261124';parent=base/'Cold713SmallLoadContinuation_20261124'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'PeakAdjointCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
path='src/Solver/CCoupledDeformationContinuation.cls';code=(root/path).read_text('utf-8-sig')
before='StopSearch "Суммарный спад найден, но его максимум не подтверждён повтором плеч."';assert code.count(before)==1
code=code.replace(before,before+' & " Причина: " & mProbeDetails',1)
values[path]=code.replace('\n','\r\n').encode('utf-8-sig')
path='src/Solver/CCoupledNewtonSolver.cls';code=(root/path).read_text('utf-8-sig')
anchor='''    Dim adjoint As New CCoupledLinearSystem
    If Not adjoint.Solve(matrix, weights, size) Then
''';assert code.count(anchor)==1
code=code.replace(anchor,'''    Dim adjoint As New CCoupledLinearSystem
    Dim adjointIndexTest As Long, adjointSolvedTest As Boolean
    adjointIndexTest = modTestCoupledElasticState.TestRecordAdjointInput(matrix, weights, size, mTrialLambda, EquilibriumTolerance(absoluteTolerance, relativeTolerance))
    adjointSolvedTest = adjoint.Solve(matrix, weights, size)
    modTestCoupledElasticState.TestRecordAdjointOutcome adjointIndexTest, adjointSolvedTest, adjoint.MinimumPivot, adjoint.FailureCode
    If Not adjointSolvedTest Then
''',1)
values[path]=code.replace('\n','\r\n').encode('utf-8-sig')
path='tests/modTestCoupledElasticState.bas';code=(root/path).read_text('utf-8-sig')
code=code.replace('Option Explicit\n','Option Explicit\nPrivate mAdjointProbeCountTest As Long\n',1)
code+='\n'+(here/'TestAdjointProbe.bas.txt').read_text('utf-8-sig')
values[path]=code.replace('\n','\r\n').encode('utf-8-sig')
imports=['CCoupledDeformationContinuation','CCoupledNewtonSolver','modTestCoupledElasticState']
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'",'-ModuleNames @('+','.join(repr(x) for x in imports)+')')
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
for name in (Path(__file__).name,'TestAdjointProbe.bas.txt'):values[name]=(here/name).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='PeakAdjointInputProbe',actual_import_this_pass=imports,
    exact_adjoint_matrix_and_rhs_recorded=True,original_LU_and_false_return_unchanged=True,
    original713_numerical_choices_preserved=True,explicit_failure_cause_added=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('715 EXACT ADJOINT INPUT DIAGNOSIS; ORIGINAL713 NUMERICS;714 PREPARATION ASSERTION FAILED BEFORE OFFICE')
