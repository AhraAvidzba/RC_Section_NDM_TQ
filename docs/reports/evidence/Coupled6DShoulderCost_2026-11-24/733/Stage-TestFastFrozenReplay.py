"""ДЛЯ ТЕСТОВ. Проверяет замороженный путь732 одним StressOnly-проходом с наблюдением каждого принятого поля."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re,struct,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold733FastFrozenReplay_20261124';parent=base/'Cold732ProvidedJacobianCapacity_20261124'
proof=json.loads((parent/'ProvidedCapacityStopProof.json').read_bytes())
assert proof['native']['own_process_closed'] and proof['accepted_states']==71
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](parent/'Native/Stock.RawHistory.bin')
frame=frames[-1];assert len(frame['fields'])==71 and raw['incomplete_tail_bytes']==0
template=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Native/PostCrackPrescribed.Accepted2.bin')
for key in ('fields','fractions','lambdas'):assert np.array_equal(frame[key][:2],template[key]),key
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'FastReplayCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
values={};prep=json.loads((root/'Preparation.json').read_bytes())
data=bytearray(b'NDMPOSTREPLAY001'+struct.pack('<i',60))
def TestVector(v):
    """Сохраняет исходные Double без округления."""
    a=np.asarray(v,dtype='<f8');return struct.pack('<i',len(a))+a.tobytes()
for key in ('calibration','options','offset','direction'):data+=TestVector(template[key])
data+=struct.pack('<i',len(frame['fields']))
for f,l,x in zip(frame['fractions'],frame['lambdas'],frame['fields']):data+=struct.pack('<dd',f,l)+TestVector(x)
values['Input/PostCrackReplay.bin']=bytes(data)
values['Input/PostCrackReplayOptions.bin']=struct.pack('<i3d',3,0,0,0)
values['Input/PostCrackSourceRawHistory.bin']=(parent/'Native/Stock.RawHistory.bin').read_bytes()
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
method,=re.findall(r'(?ms)^Friend Function ReplayAcceptedFields\(.*?^End Function',code)
anchor='        mAcceptedRevision = mAcceptedRevision + 1\n';assert method.count(anchor)==1
changed=method.replace(anchor,anchor+'        modTestCoupledElasticState.TestObserveFrozenReplay index, field\n',1)
assert changed.replace('        modTestCoupledElasticState.TestObserveFrozenReplay index, field\n','')==method
code=code.replace(method,changed,1);values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(base/'Cold728FastShoulderWitness_20261124'/name).read_text('utf-8-sig')
anchor='Option Explicit\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'Private mFrozenReplayActiveTest As Boolean, mFrozenReplayHandleTest As Integer\n',1)
anchor='        TestExportSpatialReverses solver, operators, concrete(2), response(2), sides, calibration, options, offset, direction, fields, trace\n';assert code.count(anchor)==1
code=code.replace(anchor,'''        Set mReverseFieldsTest = fields: Set mReverseTraceTest = trace
        TestExportFastFrozenReplay
''',1)
method,=re.findall(r'(?ms)^Private Sub TestExportReverseCandidate\(.*?^End Sub',code)
header=method[:method.index('    mReverseSolverTest.ResetHistory\n')]
header=header.replace('Private Sub TestExportReverseCandidate(ByVal caseIndex As Long)',
    'Private Sub TestExportFastFrozenReplay()\n    Const caseIndex As Long = 0\n    On Error GoTo FailedReplayTest',1)
code+="\n' -------------------- ДЛЯ ТЕСТОВ --------------------\n' Восстанавливает исходную историю один раз; каждый узел проверяется по фактическим тензорам без вычисления касательной.\n"+header+'''
    mFrozenReplayHandleTest = handle: mFrozenReplayActiveTest = True
    If Not mReverseSolverTest.ReplayAcceptedFields(mReverseFieldsTest, mReverseFieldsTest.Count) Then _
        Err.Raise 5, , "Не восстановлена замороженная последовательность."
    mFrozenReplayActiveTest = False
    Close #handle
    Exit Sub
FailedReplayTest:
    Dim number As Long, source As String, description As String
    number = Err.Number: source = Err.Source: description = Err.Description
    mFrozenReplayActiveTest = False
    On Error Resume Next
    If handle > 0 Then Close #handle
    On Error GoTo 0
    Err.Raise number, source, description
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Наблюдает уже принятый материал; нулевая проба читает тот же тензор и не меняет исходные поля или историю.
Public Sub TestObserveFrozenReplay(ByVal index As Long, ByRef field() As Double)
    If Not mFrozenReplayActiveTest Then Exit Sub
    Dim factor As Double, fraction As Double, load(1 To 6) As Double, external() As Double, weak() As Double, matrix() As Double
    Dim degree As Long, norm As Double, externalNorm As Double, bound As Double
    mReverseTraceTest.ReadPoint index, fraction, factor
    For degree = 1 To 6: load(degree) = fraction * mReverseBaseTest(degree) + factor * mReverseDirectionTest(degree): Next degree
    mReverseOperatorsTest.LoadEquations load(1), load(2), load(3), load(4), load(5), load(6), external
    If Not mReverseSolverTest.TestEquilibriumLinearization(field, False, True, weak, matrix) Then Err.Raise 5, , "Нет уравнений замороженной точки."
    For degree = 1 To mReverseSolverTest.PhysicalDegreeCount
        norm = GeomMax(norm, Abs(weak(degree) - external(degree))): externalNorm = GeomMax(externalNorm, Abs(external(degree)))
    Next degree
    For degree = mReverseSolverTest.PhysicalDegreeCount + 1 To mReverseOperatorsTest.DegreeCount
        norm = GeomMax(norm, mReverseModulusTest * Abs(weak(degree) - external(degree)))
    Next degree
    bound = mReverseOptionsTest(3) + mReverseOptionsTest(4) * externalNorm
    Print #mFrozenReplayHandleTest, "F;" & CStr(index) & ";" & TestOnsetNumber(fraction) & ";" & TestOnsetNumber(factor) & TestHardPreparedVector(field)
    Print #mFrozenReplayHandleTest, "R;" & CStr(index) & ";" & TestOnsetNumber(norm) & ";" & TestOnsetNumber(bound) & ";" & TestOnsetNumber(mReverseReinforcedTest.AcceptedSteelUtilization)
    TestWriteHardPreparedStress mFrozenReplayHandleTest, index, mReverseOperatorsTest, mReverseConcreteTest, mReverseReinforcedTest, mReverseSidesTest
    If norm > bound Or mReverseReinforcedTest.AcceptedSteelUtilization > 1# Then Err.Raise 5, , "Замороженная точка не прошла исходную приёмку."
End Sub
'''
assert max(map(len,code.splitlines()))<1024
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames @('CCoupledNewtonSolver','modTestCoupledElasticState')")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values['Run-TestPostCrackRawReplay.ps1']=(base/'Cold728FastShoulderWitness_20261124/Run-TestPostCrackRawReplay.ps1').read_bytes()
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,content in values.items():(root/name).write_bytes(content);prep['inputs'][name]=sha256(content).hexdigest().upper()
prep['directed_scope'].update(case='Frozen71SinglePassStressReplay',actual_import_this_pass=['CCoupledNewtonSolver','modTestCoupledElasticState'],
    source_states=71,original_binary_fields_and_loads_preserved=True,original732_numeric_methods_exact_except_readonly_replay_observer=True,
    same_native_force_gauge_and_steel_checks=True,full_independent_field_validation_required=True,
    no_new_continuation_or_capacity_search=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('733 EXACT71 FROZEN FIELDS; ONE STRESSONLY HISTORY PASS; READONLY OBSERVER; NO NEW SEARCH')
