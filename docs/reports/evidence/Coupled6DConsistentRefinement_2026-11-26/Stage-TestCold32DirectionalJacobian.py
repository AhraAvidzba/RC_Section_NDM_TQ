"""ДЛЯ ТЕСТОВ. Читает рабочие касательные и реальные приращения невязки на точном холодном состоянии32; нового равновесия не ищет."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold895FreshSpatialArc_20261126';source=base/'Cold898FreshArcColdStress_20261126'
root=base/'Cold899Cold32DirectionalJacobian_20261126'
proof=json.loads((source/'VerifiedColdStressReplayProof.json').read_bytes())
assert proof['source_states']==32 and proof['all_three_station_steel_limits_checked']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'DirectionalJacobianCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
entry='''Public Function TestReadCold32DirectionalJacobian() As String
    ' -------------------- ДЛЯ ТЕСТОВ --------------------
    ' Сопоставляет уточнённую и обычную касательные с односторонними приращениями невязки; принятую историю не изменяет.
    Dim accepted() As Double, previous() As Double, movement() As Double, projection() As Double
    mReverseSolverTest.DiscardStep
    mReverseSolverTest.CopyAccepted accepted: previous = mReverseFieldsTest(mReverseFieldsTest.Count - 1)
    movement = accepted
    Dim i As Long, j As Long, size As Long: size = UBound(accepted) + 1
    For i = 1 To UBound(movement): movement(i) = accepted(i) - previous(i): Next i
    Dim coordinate As Double, spacing As Double
    If Not mReverseSolverTest.SelectSpatialDirectionControl(previous, accepted, movement, projection, coordinate, spacing) Then _
        Err.Raise 5, , "Не определена мера принятого холодного приращения."
    Dim fraction As Double, factor As Double, previousFactor As Double
    mReverseTraceTest.ReadPoint mReverseTraceTest.Count, fraction, factor
    mReverseTraceTest.ReadPoint mReverseTraceTest.Count - 1, fraction, previousFactor
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\Cold32DirectionalJacobian.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMCOLD32JAC0001"
    Put #handle, , magic
    TestWriteReverseBinaryVector handle, Array(factor, previousFactor, revision, mReverseModulusTest, size, coordinate, spacing)
    TestWriteReverseBinaryVector handle, accepted
    TestWriteReverseBinaryVector handle, movement
    TestWriteReverseBinaryVector handle, projection
    Dim level As Variant, delta As Variant, seed() As Double, seedFactor As Double
    Dim zero() As Double, residual() As Double, matrix() As Double, ordinary() As Double, plain() As Double
    Dim probe() As Double, values() As Double, unused() As Double, scales As Variant
    Dim refinedReady As Boolean, ordinaryReady As Boolean, ready As Boolean
    Dim column() As Double: ReDim column(1 To size)
    Dim equations As ICoupledEquationSystem: Set equations = mReverseSolverTest
    For Each level In Array(0#, 0.015625, 1#, 8#)
        seed = accepted
        For i = 1 To UBound(seed): seed(i) = accepted(i) + CDbl(level) * movement(i): Next i
        seedFactor = factor + CDbl(level) * (factor - previousFactor)
        ReDim zero(1 To size)
        If Not mReverseSolverTest.BeginNewtonLineTest(mReverseBaseTest, mReverseDirectionTest, projection, _
                coordinate + CDbl(level) * spacing, seed, seedFactor) Then Err.Raise 5, , "Не подготовлен читающий контекст холодной производной."
        scales = mReverseSolverTest.ReadNewtonLineScalesTest
        refinedReady = mReverseSolverTest.TryReadBranchLinearization(zero, residual, matrix)
        ordinaryReady = equations.TryEvaluate(zero, True, plain, ordinary)
        TestWriteReverseBinaryVector handle, Array(CDbl(level), seedFactor, Abs(CLng(refinedReady)), Abs(CLng(ordinaryReady)))
        TestWriteReverseBinaryVector handle, seed
        TestWriteReverseBinaryVector handle, scales
        If Not refinedReady Or Not ordinaryReady Then Err.Raise 5, , "Не прочитаны обе касательные; неудача не является разрушением."
        TestWriteReverseBinaryVector handle, residual
        TestWriteReverseBinaryVector handle, plain
        For j = 1 To size
            For i = 1 To size: column(i) = matrix(j, i): Next i
            TestWriteReverseBinaryVector handle, column
        Next j
        For j = 1 To size
            For i = 1 To size: column(i) = ordinary(j, i): Next i
            TestWriteReverseBinaryVector handle, column
        Next j
        For Each delta In Array(-8#, -1#, -0.125, -0.015625, 0.015625, 0.125, 1#, 8#)
            ReDim probe(1 To size)
            For i = 1 To UBound(movement): probe(i) = CDbl(delta) * movement(i) / scales(i): Next i
            probe(size) = CDbl(delta) * (factor - previousFactor) / scales(size)
            ready = equations.TryEvaluate(probe, False, values, unused)
            TestWriteReverseBinaryVector handle, Array(CDbl(delta), Abs(CLng(ready)))
            TestWriteReverseBinaryVector handle, probe
            If Not ready Then Err.Raise 5, , "Не прочитан материальный отклик односторонней пробы."
            TestWriteReverseBinaryVector handle, values
        Next delta
        mReverseSolverTest.EndNewtonLineTest
        If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Чтение касательной изменило принятую ревизию."
    Next level
    Close #handle
    Dim checked() As Double: mReverseSolverTest.CopyAccepted checked
    For i = 1 To UBound(accepted)
        If checked(i) <> accepted(i) Then Err.Raise 5, , "Чтение касательной изменило принятое поле."
    Next i
    TestReadCold32DirectionalJacobian = "TOTAL_COLD32_DIRECTIONAL_JACOBIAN_TEST: passed=1; failed=0; read-only; discrepancies recorded; not capacity"
End Function
'''
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='Private Const HARD_PREPARED_HALF_TEST As Double = 0.001\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+"' FOR TESTS: baseline entry point only; existing methods retained.\n\n\n"+entry,1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'modTestCoupledElasticState' -EntryPointNames 'TestReadCold32DirectionalJacobian'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1'
runner=(base/'Cold857SpatialDeclineRecovery_20261125'/name).read_text('utf-8-sig')
anchor='        else { $result = [string]$excel.Run($macroName) }';assert runner.count(anchor)==1
runner=runner.replace(anchor,anchor+'''
        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected cold32 derivative setup macro.'}
        $lines.Add([string]$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestReadCold32DirectionalJacobian"))
        Save-Progress
''',1)
values[name]=runner.replace('\n','\r\n').encode('utf-8-sig')
values.update({'Input/Reverse1.bin':(source/'Input/Reverse1.bin').read_bytes(),
 'SourceStress32Proof.json':(source/'VerifiedColdStressReplayProof.json').read_bytes(),
 'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():
    if name.endswith('.bas'):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='TrueCold32ReadOnlyDirectionalJacobian',actual_import_this_pass=['modTestCoupledElasticState'],
 actual_component_count=256,all_working_sources_exact895=True,source_states=32,true_cold897_history=True,
 no_new_equilibrium_or_capacity_search=True,no_material_acceptance=True,levels=[0,1/64,1,8],
 signed_directional_steps=[-8,-1,-1/8,-1/64,1/64,1/8,1,8],actual_failed_final_iterate_not_captured=True,
 physical_capacity_certified=False,no_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'COLD32 EXACT;ONE SHORT TEST ENTRY;NO WORKING CHANGE;DIRECTIONAL DERIVATIVE AUDIT')
