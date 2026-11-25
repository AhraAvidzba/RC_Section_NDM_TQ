"""ДЛЯ ТЕСТОВ. Читающий захват последнего направления Ньютона и ответов вдоль него на фиксированном холодном конце."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold804HistoryAccuracyEndpoint_20261125';root=base/'Cold805EndpointLineWitness_20261125'
proof=json.loads((parent/'ColdEndpointArcProof.json').read_bytes())
assert proof['native']['own_process_closed'] and not proof['returned_verified_candidate']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'EndpointLineCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
code=code.replace('Option Explicit\n','Option Explicit\nPrivate mEndpointLineEnabledTest As Boolean, mEndpointLineTest As Variant\n',1)
method,=re.findall(r'(?ms)^Private Function SolveStep\(.*?^End Function',code)
anchor='        ReDim mCandidate(1 To mDegreeCount)\n';assert method.count(anchor)==1
code=code.replace(method,method.replace(anchor,
 '        If mEndpointLineEnabledTest Then TestRememberEndpointLine current, currentLambda, correction, norm, tolerance\n'+anchor,1),1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Включает только копирование последней поправки, без вычисления материала и изменения рабочего бюджета.
Friend Sub TestBeginEndpointLineObserver()
    mEndpointLineTest = Empty: mEndpointLineEnabledTest = True
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Отключает читающее наблюдение и возвращает копии фактического линейного направления.
Friend Function TestEndEndpointLineObserver() As Variant
    mEndpointLineEnabledTest = False: TestEndEndpointLineObserver = mEndpointLineTest
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Копирует матрицу до проб вдоль направления; никакая материальная история не принимается.
Private Sub TestRememberEndpointLine(ByRef field() As Double, ByVal factor As Double, ByRef correction() As Double, _
        ByVal norm As Double, ByVal tolerance As Double)
    Dim flat() As Double, i As Long, j As Long, k As Long
    ReDim flat(1 To mDegreeCount * mDegreeCount)
    For i = 1 To mDegreeCount
        For j = 1 To mDegreeCount: k = k + 1: flat(k) = mMatrix(i, j): Next j
    Next i
    mEndpointLineTest = Array(Array(factor, mTargetControl, mArcLoadProjection, norm, tolerance, _
        mAcceptedRevision, mIterations, mDegreeCount, mGaugeStart - 1, ControlResidual(field)), _
        field, correction, mResidual, flat, mControlProjection, mDirectionEquations)
End Sub
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
method,=re.findall(r'(?ms)^Private Sub TestRunOperationalArc\(\).*?^End Sub',code)
anchor='    solved = arc.TryStep(mReverseFieldsTest, mReverseTraceTest, ratioData(1), CLng(allowanceData(1)), factor, projection)\n'
assert method.count(anchor)==1
updated=method.replace(anchor,'    mReverseSolverTest.TestBeginEndpointLineObserver\n'+anchor+
 '    Dim lineWitness As Variant: lineWitness = mReverseSolverTest.TestEndEndpointLineObserver\n',1)
updated=updated.replace('\nEnd Sub','\n    TestExportEndpointLineWitness lineWitness\nEnd Sub',1)
code=code.replace(method,updated,1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' После завершения исходной попытки записывает её матрицу и повторяет ответы материала вдоль фактической поправки.
Private Sub TestExportEndpointLineWitness(ByRef witness As Variant)
    If IsEmpty(witness) Then Err.Raise 5, , "Нет наблюдения последнего направления Ньютона."
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\EndpointLine.Result.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMENDLINE000001"
    Put #handle, , magic
    Dim vector As Variant
    For Each vector In witness: TestWriteReverseBinaryVector handle, vector: Next vector
    Dim fractions As Variant: fractions = Array(-1#, 0#, 0.000001, 0.0001, 0.01, 0.1, 1#)
    TestWriteReverseBinaryVector handle, fractions
    Close #handle
    Dim context As Variant, original() As Double, correction() As Double, field() As Double
    context = witness(0): original = witness(1): correction = witness(2)
    Dim index As Long, degree As Long, factor As Double
    For index = LBound(fractions) To UBound(fractions)
        field = original
        For degree = 1 To UBound(field): field(degree) = original(degree) + fractions(index) * correction(degree): Next degree
        factor = context(0) + fractions(index) * correction(UBound(correction))
        TestExportSignedArcTrial "EndpointLine.Sample" & CStr(index + 1), field, factor
    Next index
End Sub
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
imports=['CCoupledNewtonSolver','modTestCoupledElasticState']
values.update({'Compile-TestExternalPoint.ps1':compiler.replace("-ModuleNames 'modTestCoupledElasticState'",
 '-ModuleNames @('+','.join(repr(x) for x in imports)+')',1).replace('\n','\r\n').encode('utf-8-sig'),
 'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():
 if n.endswith(('.cls','.bas')):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
 (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ExactEndpointNewtonDirectionWitness',actual_import_this_pass=imports,actual_component_count=256,
 source45fields_and_materials_exact804=True,working_numerical_operations_unchanged=True,
 one_read_only_hook_after_correction_before_line_search=True,
 tangent_residual_and_direction_copied_without_evaluation=True,
 seven_post_attempt_material_replays_not_counted_as_original_solver_work=True,
 negative_samples_not_equilibria_and_not_accepted=True,no_new_capacity_search=True,
 physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('805 READONLY LAST NEWTON DIRECTION; SAME804 ARC WORK;7 POST-ATTEMPT MATERIAL WITNESSES')
