"""ДЛЯ ТЕСТОВ. Читает действительные нормированные матрицы из точной принятой истории без нового поиска."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold736FrozenTailMatrix_20261125';parent=base/'Cold735FrozenGrowthReplay_20261125'
proof=json.loads((parent/'FrozenCapacityReplayProof.json').read_bytes());assert proof['all_fields_independent_weak_gauge_and_six_force_checks']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'TailMatrixCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
values={};prep=json.loads((root/'Preparation.json').read_bytes())
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
method,=re.findall(r'(?ms)^Public Sub TestObserveFrozenReplay\(.*?^End Sub',code)
anchor='    If norm > bound Or mReverseReinforcedTest.AcceptedSteelUtilization > 1# Then Err.Raise 5, , "Замороженная точка не прошла исходную приёмку."\n'
assert method.count(anchor)==1
code=code.replace(method,method.replace(anchor,anchor+'    If index = mReverseFieldsTest.Count Then TestWriteFrozenTailMatrices field, factor\n',1),1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет исходную и две направленные линеаризации от одной принятой истории; не решает пробу и не меняет материалы.
Private Sub TestWriteFrozenTailMatrices(ByRef accepted() As Double, ByVal factor As Double)
    Dim previous() As Double, seed() As Double, fraction As Double, priorFactor As Double
    previous = mReverseFieldsTest(mReverseFieldsTest.Count - 1)
    mReverseTraceTest.ReadPoint mReverseFieldsTest.Count - 1, fraction, priorFactor
    Dim mode As Long, ratio As Double, degree As Long, seedFactor As Double
    Dim projection() As Double: ReDim projection(1 To mReverseOperatorsTest.DegreeCount): projection(1) = 1#
    For mode = 0 To 2
        ratio = 0#: If mode = 1 Then ratio = 0.25
        If mode = 2 Then ratio = 1.25
        seed = accepted
        For degree = 1 To mReverseSolverTest.PhysicalDegreeCount
            seed(degree) = accepted(degree) + ratio * (accepted(degree) - previous(degree))
        Next degree
        seedFactor = factor + ratio * (factor - priorFactor)
        If Not mReverseSolverTest.BeginNewtonLineTest(mReverseBaseTest, mReverseDirectionTest, projection, seed(1), seed, seedFactor) Then _
            Err.Raise 5, , "Не открыт контекст контрольной матрицы."
        Dim equations As ICoupledEquationSystem: Set equations = mReverseSolverTest
        Dim zero() As Double, residual() As Double, matrix() As Double, scales As Variant, ready As Boolean
        ReDim zero(1 To mReverseOperatorsTest.DegreeCount + 1)
        ready = equations.TryEvaluate(zero, True, residual, matrix)
        scales = mReverseSolverTest.ReadNewtonLineScalesTest
        Dim handle As Integer: handle = FreeFile
        Open mPolishRootTest & "\\Native\\TailMatrix" & CStr(mode) & ".bin" For Binary Access Write As #handle
        Dim magic As String * 16: magic = "NDMTAILMATRIX001"
        Dim size As Long: size = UBound(zero)
        Dim flag As Long: flag = Abs(CLng(ready))
        Put #handle, , magic: Put #handle, , size: Put #handle, , flag
        Put #handle, , ratio: Put #handle, , seedFactor
        TestWriteReverseBinaryVector handle, seed
        TestWriteReverseBinaryVector handle, scales
        If ready Then
            TestWriteReverseBinaryVector handle, residual
            Dim row As Long, column As Long, value As Double
            For row = 1 To size
                For column = 1 To size
                    value = matrix(column, row): Put #handle, , value
                Next column
            Next row
        End If
        Close #handle
        mReverseSolverTest.EndNewtonLineTest
    Next mode
End Sub
'''
assert max(map(len,code.splitlines()))<1024
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_bytes();values['Compile-TestExternalPoint.ps1']=compiler
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,b in values.items():(root/name).write_bytes(b);prep['inputs'][name]=sha256(b).hexdigest().upper()
prep['directed_scope'].update(case='FrozenTailNativeScaledJacobians',actual_import_this_pass=['modTestCoupledElasticState'],
    source_root=proof['source_root'],source_replay_root=str(parent),source_states=proof['source_states'],
    native_numeric_kernel_exact_to_parent=True,no_new_capacity_search=True,read_only_tail_linearizations=[0,.25,1.25],
    matrix_format='native normalized controlled equations, row-major Double; full61x61 including lambda/control and gauges',
    physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('736 READONLY NATIVE CONTROLLED MATRICES STAGED; ALL CORE SOURCES EXACT735')
