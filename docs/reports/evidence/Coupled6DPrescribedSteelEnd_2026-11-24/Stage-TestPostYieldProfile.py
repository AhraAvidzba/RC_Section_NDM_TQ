"""ДЛЯ ТЕСТОВ. Изолирует следующий шаг от завершённого префикса723 и измеряет фазы Ньютона без изменения формул."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json, struct, re
import numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold726PostYieldProfile_20261124';parent=base/'Cold723PriorBranchCapacity_20261124'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](parent,'PostCrackTension')
assert native['timeout600'] and native['own_process_closed']
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](parent/'Native/Stock.RawHistory.bin')
frame=frames[-1];assert len(frame['fields'])==53 and raw['incomplete_tail_bytes']==0
template=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Native/PostCrackPrescribed.Accepted2.bin')
for key in ('fields','fractions','lambdas'):assert np.array_equal(frame[key][:2],template[key]),key
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ProfileCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
values={};prep=json.loads((root/'Preparation.json').read_bytes())
data=bytearray(b'NDMPOSTREPLAY001'+struct.pack('<i',60))
def TestVector(values):
    """Сохраняет исходные Double без текстового округления."""
    values=np.asarray(values,dtype='<f8');return struct.pack('<i',len(values))+values.tobytes()
for key in ('calibration','options','offset','direction'):data+=TestVector(template[key])
data+=struct.pack('<i',len(frame['fields']))
for fraction,factor,field in zip(frame['fractions'],frame['lambdas'],frame['fields']):data+=struct.pack('<dd',fraction,factor)+TestVector(field)
target=float(frame['fields'][-1,0]+1.25*(frame['fields'][-1,0]-frame['fields'][-2,0]))
values['Input/PostCrackReplay.bin']=bytes(data)
values['Input/PostCrackReplayOptions.bin']=struct.pack('<i3d',3,1,target,30)
values['Input/PostCrackSourceRawHistory.bin']=(parent/'Native/Stock.RawHistory.bin').read_bytes()
name='tests/modTestCoupledElasticState.bas'
code=(base/'Cold720PostCrackRawReplay_20261124'/name).read_text('utf-8-sig')
anchor='Option Explicit\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'''Private mReplayMaterialsTest As Collection
Private mReplayProfileRootTest As String, mReplayStartedTest As Double
''',1)
anchor='        Dim solved As Boolean, actual() As Double\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'''        Set mReplayMaterialsTest = New Collection
        For index = LBound(concrete) To UBound(concrete): mReplayMaterialsTest.Add concrete(index): Next index
        mReplayProfileRootTest = evidenceRoot: mReplayStartedTest = Timer
        TestRecordReplayPhase "DIRECT_BEGIN", 0, factor, 0#
''',1)
anchor='        If solver.AcceptedRevision <> revision Then Err.Raise 5, , "Прямая диагностическая проба изменила принятую историю."';assert code.count(anchor)==1
code=code.replace(anchor,'''        TestRecordReplayPhase "DIRECT_END", -1, trialFactor, 0#
        Set mReplayMaterialsTest = Nothing
'''+anchor,1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет стоимость фаз и локальные счётчики без изменения проб, полей и истории.
Public Sub TestRecordReplayPhase(ByVal stage As String, ByVal iteration As Long, ByVal factor As Double, ByVal norm As Double)
    If mReplayMaterialsTest Is Nothing Then Exit Sub
    Dim values As Variant, totals(0 To 4) As Double, side As Long, index As Long
    Dim material As CCoupledConcreteResponse
    For side = 1 To mReplayMaterialsTest.Count
        Set material = mReplayMaterialsTest(side): values = material.ReadDilationWorkTest
        For index = 0 To 4: totals(index) = totals(index) + CDbl(values(index)): Next index
    Next side
    Dim elapsed As Double: elapsed = Timer - mReplayStartedTest: If elapsed < 0# Then elapsed = elapsed + 86400#
    Dim handle As Integer: handle = FreeFile
    Open mReplayProfileRootTest & "\\Native\\ReplayProfile.csv" For Append As #handle
    Print #handle, stage & ";" & CStr(iteration) & ";" & TestOnsetNumber(elapsed) & ";" & TestOnsetNumber(factor) & ";" & TestOnsetNumber(norm) & TestHardPreparedVector(totals)
    Close #handle
End Sub
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
method,=re.findall(r'(?ms)^Private Function SolveStep\(.*?^End Function',code);changed=method
def TestBefore(anchor,stage,factor='currentLambda',norm='norm'):
    """Вставляет единственную строку диагностики; удаление этих строк точно восстанавливает исходный метод."""
    global changed
    assert changed.count(anchor)==1,anchor
    changed=changed.replace(anchor,'        modTestCoupledElasticState.TestRecordReplayPhase "'+stage+'", mIterations, '+factor+', '+norm+'\n'+anchor,1)
TestBefore('        built = Evaluate(current, True, useGradientTangent, Not (mConsistentControl Or mRefinedControlTangent), _','TANGENT_BEGIN')
TestBefore('        If mConsistentControl And norm > tolerance And norm <= SMALL_STRAIN_MAX_NORM * tolerance Then','TANGENT_END')
TestBefore('        If Not SolveCorrection(current, correction) Then','LINEAR_BEGIN')
TestBefore('        mMinimumPivot = mSystem.MinimumPivot\n        If useSteelCorrection Then','LINEAR_END')
TestBefore('            If Evaluate(mCandidate, False) Then','LINE_SEARCH_BEGIN','candidateLambda')
TestBefore('            If accepted Then Exit Do','LINE_SEARCH_END','candidateLambda','candidateNorm')
TestBefore('            If Not SolveStep And allowHybridRecovery Then SolveStep = TryHybridControlledField','HYBRID_BEGIN')
restored=re.sub(r'(?m)^        modTestCoupledElasticState\.TestRecordReplayPhase[^\n]*\n','',changed)
assert restored==method
code=code.replace(method,changed,1);values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames @('CCoupledNewtonSolver','modTestCoupledElasticState')")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
runner=(base/'Cold720PostCrackRawReplay_20261124/Run-TestPostCrackRawReplay.ps1').read_text('utf-8-sig').replace('-TimeoutSeconds 600','-TimeoutSeconds 180')
values['Run-TestPostCrackRawReplay.ps1']=runner.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,content in values.items():
    assert not name.endswith(('.cls','.bas')) or max(map(len,content.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(content);prep['inputs'][name]=sha256(content).hexdigest().upper()
prep['directed_scope'].update(case='OnePostYieldStepWithPhaseProfile',actual_import_this_pass=['CCoupledNewtonSolver','modTestCoupledElasticState'],
    source_states=53,target_strain=target,maximum_iterations=30,original_material722_exact=True,
    Newton_numerical_method_exact_after_removing_test_logging=True,all_source_fields_and_loads_bits_exact=True,
    previous_cold723_timeout600=True,no_new_cold_or_capacity_search=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('726 ONE DIRECT STEP FROM EXACT53 FIELDS',target,'READ-ONLY PHASE COUNTERS; SAME NUMERICAL METHOD')
