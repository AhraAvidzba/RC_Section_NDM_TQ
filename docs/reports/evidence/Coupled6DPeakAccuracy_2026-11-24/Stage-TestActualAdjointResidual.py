"""ДЛЯ ТЕСТОВ. Записывает фактическую компонентную невязку и нативные сопряжённые веса, не меняя оценку погрешности715."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold716ActualAdjointResidual_20261124';parent=base/'Cold715PeakAdjointProbe_20261124'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ActualAdjointCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
path='src/Solver/CCoupledNewtonSolver.cls';code=(root/path).read_text('utf-8-sig')
anchor='adjointIndexTest = modTestCoupledElasticState.TestRecordAdjointInput(matrix, weights, size, mTrialLambda, EquilibriumTolerance(absoluteTolerance, relativeTolerance))'
assert code.count(anchor)==1
code=code.replace(anchor,anchor[:-1]+', mResidual, mGaugeStart, mReferenceModulus, ControlResidual(mCandidate))',1)
anchor='    modTestCoupledElasticState.TestRecordAdjointOutcome adjointIndexTest, adjointSolvedTest, adjoint.MinimumPivot, adjoint.FailureCode\n'
assert code.count(anchor)==1
code=code.replace(anchor,anchor+'    If adjointSolvedTest Then modTestCoupledElasticState.TestRecordAdjointWeights adjointIndexTest, weights\n',1)
values[path]=code.replace('\n','\r\n').encode('utf-8-sig')
path='tests/modTestCoupledElasticState.bas';code=(root/path).read_text('utf-8-sig')
anchor='        ByVal factor As Double, ByVal tolerance As Double) As Long\n';assert code.count(anchor)==1
code=code.replace(anchor,'''        ByVal factor As Double, ByVal tolerance As Double, ByRef residual() As Double, _
        ByVal gaugeStart As Long, ByVal referenceModulus As Double, ByVal controlDefect As Double) As Long
''',1)
anchor='signature = "NDMADJOINT0000001"';assert code.count(anchor)==1
code=code.replace(anchor,'signature = "NDMADJOINTRES001"',1)
anchor='    For row = 1 To size: Put #handle, , rightHandSide(row): Next row\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'''    Dim actual As Double
    For row = 1 To size - 1
        actual = residual(row)
        If row >= gaugeStart Then actual = actual * referenceModulus
        Put #handle, , actual
    Next row
    Put #handle, , controlDefect
''',1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет фактическое решение штатной сопряжённой LU-системы без изменения весов.
Public Sub TestRecordAdjointWeights(ByVal index As Long, ByRef weights() As Double)
    If index = 0 Then Exit Sub
    Dim handle As Integer: handle = FreeFile
    Dim size As Long: size = UBound(weights)
    Dim row As Long
    Open mStockProbeRootTest & "\\Native\\Adjoint.Weights." & CStr(index) & ".bin" For Binary Access Write As #handle
    Put #handle, , size
    For row = 1 To size: Put #handle, , weights(row): Next row
    Close #handle
End Sub
'''
values[path]=code.replace('\n','\r\n').encode('utf-8-sig')
imports=['CCoupledNewtonSolver','modTestCoupledElasticState']
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'",'-ModuleNames @('+','.join(repr(x) for x in imports)+')')
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='ActualAdjointResidual',actual_import_this_pass=imports,
    actual_scaled_residual_and_LU_weights_recorded=True,original_uniform_uncertainty_formula_unchanged=True,
    original715_numerical_choices_preserved=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('716 ACTUAL RESIDUAL AND ADJOINT WEIGHTS; ORIGINAL UNCERTAINTY AND SEARCH UNCHANGED')
