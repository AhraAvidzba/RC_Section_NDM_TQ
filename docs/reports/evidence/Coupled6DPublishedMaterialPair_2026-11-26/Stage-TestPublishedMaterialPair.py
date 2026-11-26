"""ДЛЯ ТЕСТОВ. Дополняет нативную материальную пару штатной публикацией без движения и записью неопределённости lambda."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold975ResolvedOutgoingPair_20261126';root=base/'Cold976PublishedMaterialPair_20261126'
proof=json.loads((parent/'MidpointMaterialPairProof.json').read_bytes());assert proof['pair']['passes_existing_shape_and_resolution_limits']
code=(parent/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig')
entry=re.search(r'(?ms)^Public Function TestRunMaterialCoordinatePair\(.*?^End Function\b',code).group()
entry=entry.replace('TestRunMaterialCoordinatePair','TestRunPublishedMaterialPair').replace('TOTAL_MATERIAL_MIDPOINT_PAIR_TEST','TOTAL_PUBLISHED_MATERIAL_PAIR_TEST')

def TestReplace(old,new):
    """Добавляет один проверяемый фрагмент в новый вход; исходный вход материальной пары остаётся неизменным."""
    global entry
    assert entry.count(old)==1,old[:100];entry=entry.replace(old,new,1)

TestReplace("    ' Общие30 итераций делят полный шаг и его половина; найденные поля не принимаются, критерии проверяются независимо по экспорту.",
    "    ' Общие30 итераций включают публикацию кандидатов; публикация не двигает поля, неопределённость lambda записывается штатно, история не принимается.")
TestReplace('    If mReverseFieldsTest.Count <> 25 Then',
    '    Dim savedProbeRoot As String: savedProbeRoot = mStockProbeRootTest\n    On Error GoTo TestPublicationFailure\n    If mReverseFieldsTest.Count <> 25 Then')
anchor='    Dim total As Long, budget As Long, startFactor As Double, fullFactor As Double, fullField() As Double, fullReady As Boolean'
TestReplace(anchor,anchor+'''
    Dim publicationHandle As Integer: publicationHandle = FreeFile
    Open mPolishRootTest & "\\Native\\MaterialPublication.bin" For Binary Access Write As #publicationHandle
    Dim publicationMagic As String * 16: publicationMagic = "NDMMATPUBLISH001"
    Put #publicationHandle, , publicationMagic
    TestWriteReverseBinaryVector publicationHandle, Array(reference, revision, CDbl(meta(9)))
    mAdjointProbeCountTest = 0
    Dim publicationProjection() As Double, publicationTarget As Double, publicationField() As Double, publicationActual() As Double
    Dim publicationSeed As Variant, publicationLambda As Double, publicationReady As Boolean, uncertaintyReady As Boolean
    Dim publicationEligible As Boolean, uncertainty As Double, publicationBefore As Variant, publicationAfter As Variant
    Dim publicationAdjoint As Long, publicationCandidateAfter As Boolean, publicationRemaining As Long''')
TestReplace('scalar.TryEvaluate(mapped, False, value, gradient)','scalar.TryEvaluate(mapped, True, value, gradient)')
anchor='        mReverseSolverTest.EndNewtonLineTest'
pub='''        publicationEligible = (solved And materialReady And mappedReady And norm <= tolerance And Abs(value) <= 0.000000000002)
        publicationTarget = 0#: publicationReady = False: uncertaintyReady = False: uncertainty = 0#: publicationAdjoint = 0
        publicationCandidateAfter = False: publicationLambda = factor
        ReDim publicationProjection(1 To 1): ReDim publicationField(1 To 1): ReDim publicationActual(1 To 1)
        publicationBefore = Array(0#, 0#, 0#, 0#, 0#, 0#, revision, 0#): publicationAfter = publicationBefore
        If publicationEligible Then
            ReDim publicationProjection(1 To UBound(field))
            For i = 1 To mReverseSolverTest.PhysicalDegreeCount
                publicationProjection(i) = gradient(i) / (CDbl(scales(i)) * CDbl(meta(9))) ' Материальная производная переводится из масштабированных координат в ту же физическую строку управления.
                publicationTarget = publicationTarget + publicationProjection(i) * field(i)
            Next i
        End If
        mReverseSolverTest.EndNewtonLineTest
        publicationRemaining = CLng(meta(8)) - total
        If publicationEligible And publicationRemaining > 0 Then
            publicationSeed = field
            publicationReady = mReverseSolverTest.TryDeformationStep(mReverseBaseTest, mReverseDirectionTest, publicationProjection, _
                publicationTarget, factor, publicationLambda, CDbl(meta(6)), CDbl(meta(7)), publicationRemaining, _
                consistentControl:=True, initialField:=publicationSeed)
            publicationBefore = mReverseSolverTest.ReadPolishEffortTest
            If publicationReady Then
                If publicationBefore(0) <> 0 Or publicationLambda <> factor Then Err.Raise 5, , "Штатная публикация изменила материальный корень."
                mReverseSolverTest.CopyTrial publicationField, publicationActual
                For i = 1 To UBound(field)
                    If publicationField(i) <> field(i) Then Err.Raise 5, , "Штатная публикация изменила поле материального корня."
                Next i
                mStockProbeRootTest = mPolishRootTest
                uncertaintyReady = mReverseSolverTest.TrialLambdaUncertainty(CDbl(meta(6)), CDbl(meta(7)), uncertainty)
                mStockProbeRootTest = savedProbeRoot: publicationAdjoint = mAdjointProbeCountTest
                publicationAfter = mReverseSolverTest.ReadPolishEffortTest
                publicationCandidateAfter = CBool(publicationAfter(5))
                If publicationCandidateAfter Then
                    mReverseSolverTest.CopyTrial publicationField, publicationActual
                    For i = 1 To UBound(field)
                        If publicationField(i) <> field(i) Then Err.Raise 5, , "Оценка неопределённости изменила поле материального корня."
                    Next i
                End If
            Else
                publicationAfter = publicationBefore
            End If
            total = total + CLng(publicationAfter(0))
        End If
        If total > CLng(meta(8)) Then Err.Raise 5, , "Материальная пара с публикацией превысила общий бюджет."
        TestWriteReverseBinaryVector publicationHandle, Array(caseIndex, Abs(CLng(publicationEligible)), Abs(CLng(publicationReady)), _
            publicationLambda, Abs(CLng(uncertaintyReady)), uncertainty, total, revision, mReverseSolverTest.AcceptedRevision, _
            Abs(CLng(publicationCandidateAfter)), publicationAdjoint)
        TestWriteReverseBinaryVector publicationHandle, publicationProjection
        TestWriteReverseBinaryVector publicationHandle, publicationField
        TestWriteReverseBinaryVector publicationHandle, publicationActual
        TestWriteReverseBinaryVector publicationHandle, publicationBefore
        TestWriteReverseBinaryVector publicationHandle, publicationAfter
        mReverseSolverTest.DiscardStep'''
TestReplace(anchor,pub)
TestReplace('    Close #handle\n    mReverseSolverTest.CopyAccepted checked',
'''    Close #handle: handle = 0
    TestWriteReverseBinaryVector publicationHandle, Array(0#, total, mReverseSolverTest.AcceptedRevision)
    Close #publicationHandle: publicationHandle = 0
    mStockProbeRootTest = savedProbeRoot
    mReverseSolverTest.CopyAccepted checked''')
anchor='    TestRunPublishedMaterialPair = "TOTAL_PUBLISHED_MATERIAL_PAIR_TEST: passed=1; failed=0; numerical outcomes recorded; no capacity claim"'
TestReplace(anchor,anchor+'''
    Exit Function
TestPublicationFailure:
    Dim errorNumber As Long, errorSource As String, errorText As String
    errorNumber = Err.Number: errorSource = Err.Source: errorText = Err.Description
    On Error Resume Next
    mStockProbeRootTest = savedProbeRoot
    If publicationHandle <> 0 Then Close #publicationHandle
    If handle <> 0 Then Close #handle
    mReverseSolverTest.EndNewtonLineTest: mReverseSolverTest.DiscardStep
    On Error GoTo 0
    Err.Raise errorNumber, errorSource, errorText''')
assert max(map(len,entry.splitlines()))<1024
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
ancestor=root/'PublicationPairCompilationAncestor';ancestor.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(ancestor/name)
prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n";anchor=prefix*24+'\n\nPublic Function TestRunMaterialCoordinatePair()'
assert code.count(anchor)==1
updates={'tests/modTestCoupledElasticState.bas':code.replace(anchor,prefix*25+'\n\n'+entry+'\n\n\nPublic Function TestRunMaterialCoordinatePair()',1).replace('\n','\r\n').encode('utf-8-sig'),
    'TestRunPublishedMaterialPair.bas.txt':(entry+'\n').replace('\n','\r\n').encode('utf-8-sig')}
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
updates['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'modTestCoupledElasticState' -EntryPointNames 'TestRunPublishedMaterialPair'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig');old='modTestCoupledElasticState.TestRunMaterialCoordinatePair';assert runner.count(old)==1
updates[name]=runner.replace(old,'modTestCoupledElasticState.TestRunPublishedMaterialPair',1).replace('\n','\r\n').encode('utf-8-sig')
updates.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='NativeMaterialCandidatePublicationAndLambdaUncertaintyWithoutMovement',actual_import_this_pass=['modTestCoupledElasticState'],
    actual_component_count=257,all256_numerical_components_exact975=True,only_new_test_entry_added=True,
    original_midpoint_pair_inputs_and_shared30_exact975=True,native_gradient_from_existing_material_constraint=True,
    existing_TryDeformationStep_publication_requires_zero_iterations_bitexact_field_and_lambda=True,
    existing_TrialLambdaUncertainty_and_adjoint_capture_used=True,no_material_history_acceptance=True,
    no_future_equilibrium_field_input=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'NEW TEST ENTRY ONLY;EXISTING NATIVE PUBLICATION/UNCERTAINTY;ZERO FIELD/LAMBDA MOTION;256 NUMERICAL COMPONENTS UNCHANGED')
