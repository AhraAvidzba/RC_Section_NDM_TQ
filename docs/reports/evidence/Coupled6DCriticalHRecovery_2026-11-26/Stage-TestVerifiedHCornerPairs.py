"""ДЛЯ ТЕСТОВ. Проверяет каждый рабочий H-шаг половиной от середины найденного полного поля и экспортирует обе пробы для независимого контроля."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');parent=base/'Cold874SharedCurveHControl20_20261126';root=base/'Cold876VerifiedHCornerPairs20_20261126'
proof=json.loads((parent/'SharedCurveHControlProof.json').read_bytes());assert proof['full_H_roots']==1 and proof['half_H_roots']==0
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'VerifiedHCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}

def replace(code,old,new):
    """Проверяет точный исходный контракт перед каждым изменением."""
    assert code.count(old)==1,(old[:80],code.count(old));return code.replace(old,new,1)

name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
code=replace(code,'Private mEnvelope As CCoupledEnvelopeEquations, mEnvelopeHalfConfirmed As Boolean, mEnvelopeDetails As String',
'''Private mEnvelope As CCoupledEnvelopeEquations, mEnvelopeHalfConfirmed As Boolean, mEnvelopeDetails As String
Private mEnvelopePairWitnessTest As Variant ' Читающее свидетельство пары не используется для выбора или принятия рабочих состояний.''')
code=replace(code,'    mEnvelopeHalfConfirmed = False: mEnvelopeDetails = vbNullString','    mEnvelopeHalfConfirmed = False: mEnvelopeDetails = vbNullString: mEnvelopePairWitnessTest = Empty')
code=replace(code,'    If Not SpatialIncrementAdvances(previous, mAccepted, full, advance, reference) Then\n        midpoint = mAccepted',
'''    Dim turnRequired As Boolean: turnRequired = Not SpatialIncrementAdvances(previous, mAccepted, full, advance, reference)
    Dim fullSteps As Long: fullSteps = mIterations
    midpoint = mAccepted''')
# У существующей половинной проверки убирается только внешний условный блок; все условия сил, стали и кривизны остаются.
start=code.index('    Dim turnRequired As Boolean: turnRequired = Not SpatialIncrementAdvances(')
end=code.index('    Dim movement() As Double, coordinate As Double',start)
block=code[start:end]
block=replace(block,'        mEnvelopeHalfConfirmed = True','        mEnvelopeHalfConfirmed = turnRequired')
block=replace(block,'        mEnvelopeDetails = mEnvelopeDetails & "; поворот подтверждён половинным H; ошибка поля "',
'        mEnvelopeDetails = mEnvelopeDetails & "; полная/половинная H-пара подтверждена; ошибка поля "')
assert block.endswith('    End If\n');block=block[:-len('    End If\n')]
block=replace(block,'        Dim response As CCoupledReinforcedResponse: Set response = mResponse',
'''        Dim halfActual() As Double: halfActual = mTrialResultants
        Dim response As CCoupledReinforcedResponse: Set response = mResponse''')
code=code[:start]+block+code[end:]
old='''    lambda = fullLambda
    mMeta.DiagnosticDetails = DiagnosticSummary() & " | Материальное продолжение" & mEnvelopeDetails'''
new='''    lambda = fullLambda
    If mIterations > maxIterations Then Err.Raise 5, "CCoupledNewtonSolver", "Полная и половинная H-пробы превысили общий бюджет."
    mEnvelopePairWitnessTest = Array(Array(acceptedLambda, fullLambda, halfLambda, baseline, target, 0.5 * (baseline + target), _
        fullSteps, mIterations - fullSteps, mIterations, tensorError, lambdaError, lambdaBound, jointError, curveScale, Abs(CLng(turnRequired)), _
        mAcceptedRevision, probes), mAccepted, full, half, mTrialResultants, halfActual)
    mMeta.DiagnosticDetails = DiagnosticSummary() & " | Материальное продолжение" & mEnvelopeDetails'''
code=replace(code,old,new)
code+='''

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Копирует фактическую полную и половинную H-пробы после штатной силовой приёмки; данные не возвращаются в расчёт.
Friend Function TestReadEnvelopePairVerification() As Variant
    TestReadEnvelopePairVerification = mEnvelopePairWitnessTest
End Function
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
entry='''Public Function TestRunVerifiedHPairs() As String
    ' -------------------- ДЛЯ ТЕСТОВ --------------------
    ' Проверяет три независимых общих H-шага с обязательной внутренней половиной; историю не принимает.
    Dim accepted() As Double, previous() As Double, beforeH As Double, afterH As Double, baseline As Double, resolution As Double
    Dim factor As Double, previousFactor As Double, fraction As Double, count As Long: count = mReverseFieldsTest.Count
    mReverseSolverTest.CopyAccepted accepted
    mReverseTraceTest.ReadPoint count, fraction, factor: mReverseTraceTest.ReadPoint count - 1, fraction, previousFactor
    Dim absTol As Double, relTol As Double: absTol = 0.0001 * mReverseOptionsTest(3): relTol = 0.0001 * mReverseOptionsTest(4)
    If Not mReverseSolverTest.TryEnvelopeControlState(mReverseBaseTest, mReverseDirectionTest, factor, absTol, relTol, baseline, resolution) Then Err.Raise 5, , "Нет исходного H."
    If Not mReverseSolverTest.TryAcceptedEnvelopeSegment(previous, beforeH, afterH) Then Err.Raise 5, , "Нет исходной H-пары."
    If afterH <> baseline Or afterH <= beforeH Then Err.Raise 5, , "Нет роста исходной H-пары."
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\VerifiedHPairs.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMHVERIFIED0001"
    Put #handle, , magic: TestWriteReverseBinaryVector handle, Array(count, factor, previousFactor, baseline, beforeH, resolution)
    Dim logHandle As Integer: logHandle = FreeFile
    Open mPolishRootTest & "\\Native\\VerifiedHPairs.txt" For Output As #logHandle
    Dim index As Long, i As Long, ratio As Double, target As Double, returned As Double, probes As Long, solved As Boolean
    Dim seed() As Double, projection() As Double, witness As Variant, vector As Variant
    Dim records As New Collection, record As Collection, full() As Double, half() As Double, field() As Double, actual() As Double, info As Variant
    For index = 1 To 3
        ratio = 0.25 ^ (index - 1): target = baseline + ratio * (afterH - beforeH): seed = accepted
        For i = 1 To mReverseSolverTest.PhysicalDegreeCount: seed(i) = accepted(i) + ratio * (accepted(i) - previous(i)): Next i
        solved = mReverseSolverTest.TryEnvelopeStep(mReverseBaseTest, mReverseDirectionTest, previous, factor, seed, _
            factor + ratio * (factor - previousFactor), target, returned, projection, absTol, relTol, 30, probes, False)
        witness = mReverseSolverTest.TestReadEnvelopePairVerification
        TestWriteReverseBinaryVector handle, Array(index, ratio, target, Abs(CLng(solved)), returned, probes, mReverseSolverTest.IterationCount, Abs(CLng(Not IsEmpty(witness))))
        TestWriteReverseBinaryVector handle, mReverseSolverTest.ReadPolishEffortTest
        If solved Then
            If IsEmpty(witness) Then Err.Raise 5, , "Нет свидетельства полной/половинной H-пары."
            For Each vector In witness: TestWriteReverseBinaryVector handle, vector: Next vector
            mReverseSolverTest.CopyTrial field, actual
            TestWriteReverseBinaryVector handle, field: TestWriteReverseBinaryVector handle, actual
            Set record = New Collection: record.Add index: record.Add witness: records.Add record
        ElseIf Not IsEmpty(witness) Then
            Err.Raise 5, , "Неуспешный шаг оставил пригодное свидетельство H-пары."
        End If
        Print #logHandle, "CASE;" & CStr(index) & ";SOLVED;" & CStr(solved) & ";LAMBDA;" & TestOnsetNumber(returned) & ";PRIMARY;" & CStr(mReverseSolverTest.IterationCount)
        Print #logHandle, "DETAILS;" & mReverseSolverTest.ResultMeta.DiagnosticDetails
        mReverseSolverTest.DiscardStep
    Next index
    Close #handle: Close #logHandle
    For index = 1 To records.Count
        Set record = records(index): witness = record(2): info = witness(0): full = witness(2): half = witness(3)
        TestExportSignedArcTrial "VerifiedH.Case" & CStr(record(1)) & "Full", full, CDbl(info(1))
        TestExportSignedArcTrial "VerifiedH.Case" & CStr(record(1)) & "Half", half, CDbl(info(2))
    Next index
    TestRunVerifiedHPairs = "TOTAL_VERIFIED_H_PAIRS_TEST: passed=1; failed=0; numerical outcomes recorded; no history or physical limit accepted"
End Function
'''
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig');position=code.index('Public Function TestReadSharedCurveCases()');prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n"
assert code[:position].endswith('\n\n\n');code=code[:position-2]+prefix+'\n\n'+entry+'\n\n\n'+code[position:]
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig');compiler=replace(compiler,"-ModuleNames 'modTestCoupledElasticState'","-ModuleNames @('CCoupledNewtonSolver','modTestCoupledElasticState') -EntryPointNames 'TestRunVerifiedHPairs'")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';script=(base/'Cold857SpatialDeclineRecovery_20261125'/name).read_text('utf-8-sig');anchor='        else { $result = [string]$excel.Run($macroName) }'
script=replace(script,anchor,anchor+'''
        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected verified-H setup macro.'}
        $lines.Add([string]$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestReadSharedCurveCases"))
        Save-Progress
        $lines.Add([string]$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestRunVerifiedHPairs"))
        Save-Progress
''')
values[name]=script.replace('\n','\r\n').encode('utf-8-sig')
values.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),
 'SourceSharedCurveHControlProof.json':(parent/'SharedCurveHControlProof.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='WorkingVerifiedHPairsFrom20',actual_import_this_pass=['CCoupledNewtonSolver','modTestCoupledElasticState'],actual_component_count=256,
    every_working_H_step_requires_internal_full_half_pair=True,half_seed_midpoint_of_solved_full_root=True,
    original_history_not_advanced=True,shared_curve_metric_retained=True,EnvelopeTurnVerified_preserves_actual_turn_meaning=True,
    original_pair_budget30=True,unchanged_materials_force_and_steel=True,physical_capacity_certified=False,no_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'WORKING H STEPS ALWAYS VERIFIED BY INTERNAL HALF;FULL MIDPOINT SEED;SAME30/MATERIAL/FORCE;ACTUAL TURN FLAG PRESERVED')
