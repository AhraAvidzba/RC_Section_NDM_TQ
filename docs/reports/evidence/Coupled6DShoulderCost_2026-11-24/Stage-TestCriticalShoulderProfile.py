"""ДЛЯ ТЕСТОВ. Изолирует штатное подтверждение точности первого плеча позднего максимума723."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,struct,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold727CriticalShoulderProfile_20261124';parent=base/'Cold726PostYieldProfile_20261124'
assert json.loads((parent/'DirectProbeProof.json').read_bytes())['numerical_step_solved']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'CriticalProfileCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
values={};prep=json.loads((root/'Preparation.json').read_bytes())
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](root/'Input/PostCrackSourceRawHistory.bin');frame=frames[-1]
source=(root/'Input/PostCrackReplay.bin').read_bytes()
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];reader=Reader(source)
size,=reader.ReadTest('<i');assert size==60
for _ in range(4):reader.VectorTest()
data=bytearray(source[:reader.position]+struct.pack('<i',52))
for f,l,x in zip(frame['fractions'][:52],frame['lambdas'][:52],frame['fields'][:52]):
    data+=struct.pack('<ddi',f,l,size)+np.asarray(x,dtype='<f8').tobytes()
target=float((frame['fields'][51,0]+frame['fields'][52,0])/2)
values['Input/PostCrackReplay.bin']=bytes(data)
values['Input/PostCrackReplayOptions.bin']=struct.pack('<i3d',3,1,target,30)
name='src/Solver/CCoupledDeformationContinuation.cls';code=(root/name).read_text('utf-8-sig')
anchor='        accuracyScale = accuracyScale * (0.5 * (target / point.Uncertainty))\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'        modTestCoupledElasticState.TestRecordReplayPhase "QUALITY_RETRY", attempt, accuracyScale, point.Uncertainty\n',1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет фактический ConfirmShoulderAccuracy над уже решённой пробой без принятия истории и без запуска всего поиска.
Friend Function TestConfirmComputedShoulder(ByVal solver As CCoupledNewtonSolver, ByVal reinforcement As CCoupledReinforcedResponse, _
        ByRef baseLoad() As Double, ByRef direction() As Double, ByRef projection() As Double, ByVal target As Double, ByVal factor As Double, _
        ByVal previousFactor As Double, ByVal previousCoordinate As Double, ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double) As Variant
    Set mSolver = solver: Set mReinforcement = reinforcement
    mOffset = baseLoad: mDirection = direction: mProjection = projection
    mAbsTol = absoluteTolerance: mRelTol = relativeTolerance: mMaxIterations = 30
    mLambdaTolerance = 0.001: mMaximumLambda = 3#: mProbeCount = 0: mShoulderProbes = 0
    mLocalPeak = 0.597164082727476: mFollowingLocalPeak = True
    mBranchTangentControl = False: mFieldControlChanges = 0: mUsesEnvelopeControl = False
    Set mMeta = New CResultMeta
    Dim point As TBranchPoint, initialUncertainty As Double, confirmed As Boolean
    point.Coordinate = target: point.Lambda = factor
    If mSolver.TrialLambdaUncertainty(mAbsTol, mRelTol, initialUncertainty) Then
        confirmed = ConfirmShoulderAccuracy(point, previousFactor, previousCoordinate, False)
    End If
    TestConfirmComputedShoulder = Array(Abs(CLng(confirmed)), point.Lambda, initialUncertainty, point.Uncertainty, mProbeCount, mShoulderProbes, mProbeDetails)
End Function
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='        TestRecordReplayPhase "DIRECT_END", -1, trialFactor, 0#\n';assert code.count(anchor)==1
code=code.replace(anchor,'''        Dim directSolvedTest As Boolean: directSolvedTest = solved
        Dim qualityInfoTest As String
        If solved Then
            Dim qualityBranchTest As New CCoupledDeformationContinuation, qualityResultTest As Variant, acceptedFactorTest As Double
            trace.ReadPoint fields.Count, fraction, acceptedFactorTest
            qualityResultTest = qualityBranchTest.TestConfirmComputedShoulder(solver, response(2), offset, direction, projection, target, _
                trialFactor, acceptedFactorTest, accepted(1), CDbl(options(3)), CDbl(options(4)))
            solved = (qualityResultTest(0) <> 0): trialFactor = qualityResultTest(1)
            qualityInfoTest = "QUALITY_DETAILS: initialUncertainty=" & TestOnsetNumber(qualityResultTest(2)) & _
                "; finalUncertainty=" & TestOnsetNumber(qualityResultTest(3)) & "; additionalProbes=" & CStr(qualityResultTest(4)) & _
                "; details=" & CStr(qualityResultTest(6)) & vbCrLf
        End If
'''+anchor,1)
anchor='report = "DIRECT_PROBE: solved=" & CStr(solved)';assert code.count(anchor)==1
code=code.replace(anchor,'report = qualityInfoTest & "CRITICAL_QUALITY_PROBE: directSolved=" & CStr(directSolvedTest) & "; qualityConfirmed=" & CStr(solved)',1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames @('CCoupledDeformationContinuation','modTestCoupledElasticState')")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,content in values.items():
    assert not name.endswith(('.cls','.bas')) or max(map(len,content.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(content);prep['inputs'][name]=sha256(content).hexdigest().upper()
prep['directed_scope'].update(case='ActualCriticalShoulderAccuracyProfile',actual_import_this_pass=['CCoupledDeformationContinuation','modTestCoupledElasticState'],
    source_states=52,target_strain=target,same_exact_ConfirmShoulderAccuracy=True,
    original_uniform_adjoint_estimator_and_force_acceptance=True,at_most_three_original_quality_retries=True,
    no_new_cold_capacity_search=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('727 ACTUAL CRITICAL SHOULDER ACCURACY FROM52 STATES; TARGET',target,'SAME ESTIMATOR AND RETRIES')
