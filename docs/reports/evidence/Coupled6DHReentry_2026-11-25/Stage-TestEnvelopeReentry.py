"""ДЛЯ ТЕСТОВ. Проверяет признак разгрузки113 и существующее H-продолжение от последней растущей материальной координаты, без нового закона материала."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold839ContinuedSpectralWalk_20261125';root=base/'Cold840EnvelopeReentry_20261125'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'EnvelopeReentryCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='    TestRunContinuedOutgoingGuard\n    TestRunAutomaticLoadingWalk\n    TestRecentLoadingResetGuard\n';assert code.count(anchor)==1
code=code.replace(anchor,'    TestReadUnloadingAndEnvelopeReentry\n',1)
anchor='            For principal = 1 To 3\n                available = concrete.TryReadElasticPrincipalBoundary'
assert code.count(anchor)==1
code=code.replace(anchor,'''            If point <= mReverseOperatorsTest.ConcretePointCount Then
                Dim pointEnergy As Double
                If Not steel.TryReadEnvelope(point, strain, pointEnergy) Then Err.Raise 5, , "Нет читающей материальной координаты нового корня."
                Print #handle, "E;" & CStr(caseIndex) & ";" & CStr(point) & ";" & CStr(side) & ";" & TestOnsetNumber(pointEnergy)
            End If
'''+anchor,1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает всю принятую H-историю и от последнего её роста проверяет восемь непринятых целей существующим рабочим решателем.
' Четыре размера повторяются при штатном силовом допуске и при более строгом дуговом; результаты не смешиваются.
Private Sub TestReadUnloadingAndEnvelopeReentry()
    Dim count As Long: count = mReverseFieldsTest.Count
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\EnvelopeReentry.csv" For Output As #handle
    Dim i As Long, level As Double, old As Double, lastActive As Long
    For i = 1 To count
        If Not mReverseSolverTest.TryReadAcceptedEnvelopeAt(i, count, level) Then Err.Raise 5, , "Нет полного журнала принятых материальных координат."
        Print #handle, "H;" & CStr(i) & ";" & TestOnsetNumber(level)
        If i > 1 And level > old Then lastActive = i
        old = level
    Next i
    If lastActive < 2 Or lastActive >= count Then Err.Raise 5, , "Не выделен участок с неизменной материальной координатой."
    Dim prefix As New Collection, field() As Double
    For i = 1 To lastActive: field = mReverseFieldsTest(i): prefix.Add field: Next i
    If Not mReverseSolverTest.ReplayAcceptedFields(prefix, prefix.Count) Then Err.Raise 5, , "Не восстановлена история последнего роста H."
    Dim fraction As Double, factor As Double, previousFactor As Double
    mReverseTraceTest.ReadPoint lastActive, fraction, factor
    mReverseTraceTest.ReadPoint lastActive - 1, fraction, previousFactor
    Dim baseline As Double, resolution As Double, beforeH As Double, afterH As Double, previous() As Double, accepted() As Double
    If Not mReverseSolverTest.TryEnvelopeControlState(mReverseBaseTest, mReverseDirectionTest, factor, _
            mReverseOptionsTest(3), mReverseOptionsTest(4), baseline, resolution) Then Err.Raise 5, , "Нет исходной H после восстановления."
    If Not mReverseSolverTest.TryAcceptedEnvelopeSegment(previous, beforeH, afterH) Then Err.Raise 5, , "Нет исходной растущей пары H."
    If afterH <> baseline Or afterH <= beforeH Then Err.Raise 5, , "Исходная пара H не является последним материальным ростом."
    mReverseSolverTest.CopyAccepted accepted
    Dim desired As Double: desired = GeomMax(afterH - beforeH, GeomMax(4# * resolution, 0.0001 * baseline))
    desired = GeomMin(desired, GeomMax(4# * resolution, 0.02 * baseline))
    Print #handle, "START" & TestHardPreparedVector(Array(lastActive, factor, previousFactor, baseline, beforeH, resolution, desired))
    Dim resultHandle As Integer: resultHandle = FreeFile
    Open mPolishRootTest & "\\Native\\EnvelopeReentry.bin" For Binary Access Write As #resultHandle
    Dim magic As String * 16: magic = "NDMHREENTRY000001"
    Put #resultHandle, , magic
    Dim caseIndex As Long, stepSize As Double, ratio As Double, accuracy As Double, target As Double, nextFactor As Double, probes As Long
    Dim seed() As Double, initialFactor As Double, projection() As Double, actual() As Double, solved As Boolean, energy As Double
    Dim forward As Boolean, half As Boolean, advance As Double, reference As Double, steel0 As Double, steel1 As Double, steel2 As Double
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    Dim sideSteel As CCoupledReinforcedResponse, checked() As Double
    For caseIndex = 1 To 8
        accuracy = 1#: If caseIndex > 4 Then accuracy = 0.0001
        stepSize = desired * 0.5 ^ ((caseIndex - 1) Mod 4): target = baseline + stepSize
        ratio = GeomMin(1024#, stepSize / (afterH - beforeH)): seed = accepted
        For i = 1 To mReverseSolverTest.PhysicalDegreeCount: seed(i) = accepted(i) + ratio * (accepted(i) - previous(i)): Next i
        initialFactor = factor + ratio * (factor - previousFactor)
        solved = mReverseSolverTest.TryEnvelopeStep(mReverseBaseTest, mReverseDirectionTest, previous, factor, seed, initialFactor, target, _
            nextFactor, projection, mReverseOptionsTest(3) * accuracy, mReverseOptionsTest(4) * accuracy, 30, probes, True)
        energy = 0#: advance = 0#: reference = 0#: forward = False: half = False: steel0 = 0#: steel1 = 0#: steel2 = 0#
        If solved Then
            mReverseSolverTest.CopyTrial field, actual
            If Not mReverseSolverTest.TryReadEnvelopeCoordinate(False, energy) Then Err.Raise 5, , "Нет H полностью проверенного кандидата."
            forward = mReverseSolverTest.SpatialIncrementAdvances(previous, accepted, field, advance, reference)
            half = mReverseSolverTest.EnvelopeTurnVerified
            steel0 = mReverseReinforcedTest.TrialSteelUtilization
            Set sideSteel = mReverseSidesTest(2): steel1 = sideSteel.TrialSteelUtilization
            Set sideSteel = mReverseSidesTest(4): steel2 = sideSteel.TrialSteelUtilization
            TestWriteElasticPrincipalRoot caseIndex, field
        End If
        TestWriteReverseBinaryVector resultHandle, Array(caseIndex, Abs(CLng(solved)), accuracy, baseline, target, energy, nextFactor, _
            probes, revision, mReverseSolverTest.AcceptedRevision, Abs(CLng(forward)), Abs(CLng(half)), advance, reference, steel0, steel1, steel2)
        TestWriteReverseBinaryVector resultHandle, mReverseSolverTest.ReadPolishEffortTest
        If solved Then
            TestWriteReverseBinaryVector resultHandle, field: TestWriteReverseBinaryVector resultHandle, actual
        Else
            TestWriteReverseBinaryVector resultHandle, Array(0#): TestWriteReverseBinaryVector resultHandle, Array(0#)
        End If
        Print #handle, "CASE" & TestHardPreparedVector(Array(caseIndex, Abs(CLng(solved)), accuracy, target, nextFactor, probes, energy))
        Print #handle, "DETAILS;" & CStr(caseIndex) & ";" & mReverseSolverTest.ResultMeta.DiagnosticDetails
        mReverseSolverTest.DiscardStep
        If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Проба H приняла неизвестную историю."
        mReverseSolverTest.CopyAccepted checked
        For i = 1 To UBound(accepted)
            If accepted(i) <> checked(i) Then Err.Raise 5, , "Проба H изменила исходное принятое поле."
        Next i
    Next caseIndex
    Close #resultHandle: Close #handle
    If Not mReverseSolverTest.ReplayAcceptedFields(mReverseFieldsTest, count) Then Err.Raise 5, , "Не восстановлена полная известная история после чтения проб."
End Sub
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
values.update({'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():
    if n.endswith(('.bas','.cls')):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='DetectUnloading113AndExistingEnvelopeReentry',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
    numerical_kernel_and_materials_exact839=True,last_growing_H_selected_from_accepted_journal=True,
    original_user_force_tolerance_and_separate_strict_arc_comparison=True,eight_unaccepted_trials=True,
    each_trial_shared30_iterations=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root)
