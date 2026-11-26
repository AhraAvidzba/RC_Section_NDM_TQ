"""ДЛЯ ТЕСТОВ. Переиспользует материальный тест для последовательной пары с нативным средним прогнозом и общими30 итерациями."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold973OutgoingMaterialPair_20261126';root=base/'Cold974MidpointMaterialPair_20261126'
proof=json.loads((parent/'MaterialPrefixCoordinateProof.json').read_bytes())
assert proof['cases'][0]['independently_equilibrated']
assert not proof['material_full_half_pairs'][1]['passes_independent_existing_shape_and_lambda_error_limits']
code=(parent/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig')
entry=re.search(r'(?ms)^Public Function TestSolveRecordedYieldCoordinate\(.*?^End Function\b',code).group()

def TestReplace(old,new):
    """Меняет один точный фрагмент переиспользуемого теста, не затрагивая его исходную реализацию в модуле."""
    global entry
    assert entry.count(old)==1,old[:100];entry=entry.replace(old,new,1)

entry=entry.replace('TestSolveRecordedYieldCoordinate','TestRunMaterialCoordinatePair')
entry=entry.replace('RecordedYieldCoordinate','MaterialMidpointPair').replace('NDMPOINTYIELD001','NDMMATPAIR000001')
entry=entry.replace('TOTAL_RECORDED_YIELD_COORDINATE_TEST','TOTAL_MATERIAL_MIDPOINT_PAIR_TEST')
TestReplace("    ' Решает силовые уравнения с материальной координатой вместо записанной линейной проекции, сохраняя историю25 и шесть направлений нагрузки.",
    "    ' Проверяет материальную пару от одной истории25: половина прогнозируется серединой между принятым и полным полем.")
TestReplace("    ' Найденные поля не принимаются; равновесие, поверхность и напряжения отдельно проверяются по экспорту.",
    "    ' Общие30 итераций делят полный шаг и его половина; найденные поля не принимаются, критерии проверяются независимо по экспорту.")
TestReplace('    TestWriteReverseBinaryVector handle, Array(4#)',
'''    TestWriteReverseBinaryVector handle, meta
    Dim reference As Double, unused As Double
    mReverseTraceTest.ReadPoint mReverseTraceTest.Count, unused, reference
    TestWriteReverseBinaryVector handle, Array(reference, revision)
    TestWriteReverseBinaryVector handle, accepted: TestWriteReverseBinaryVector handle, projection
    Dim total As Long, budget As Long, startFactor As Double, fullFactor As Double, fullField() As Double, fullReady As Boolean''')
TestReplace('    For caseIndex = 1 To 4','    For caseIndex = 1 To 2')
TestReplace('        target = CDbl(meta(3)): If caseIndex > 2 Then target = CDbl(meta(10))',
'''        budget = CLng(meta(8)) - total: If caseIndex = 1 Then budget = budget \\ 2
        If budget < 1 Then Exit For
        startFactor = CDbl(meta(4))
        If caseIndex = 2 Then
            If Not fullReady Then Exit For
            For i = 1 To UBound(seed): seed(i) = 0.5 * (accepted(i) + fullField(i)): Next i
            startFactor = 0.5 * (reference + fullFactor)
        End If
        target = CDbl(meta(3)): If caseIndex = 2 Then target = CDbl(meta(10))''')
TestReplace('        complete = (caseIndex Mod 2 = 1)','        complete = True')
TestReplace('                CDbl(meta(5)), seed, CDbl(meta(4)))','                CDbl(meta(5)), seed, startFactor)')
old='''        For i = 1 To UBound(seed)
            raw = values(i)
            If i > mReverseSolverTest.PhysicalDegreeCount Then raw = raw / CDbl(meta(9))
            If raw <> expected(i) Then Err.Raise 5, , "Невязка не совпала побитно с фактическим отказом."
        Next i'''
new='''        If caseIndex = 1 Then
            For i = 1 To UBound(seed)
                raw = values(i)
                If i > mReverseSolverTest.PhysicalDegreeCount Then raw = raw / CDbl(meta(9))
                If raw <> expected(i) Then Err.Raise 5, , "Невязка не совпала побитно с фактическим отказом."
            Next i
        End If
        TestWriteReverseBinaryVector handle, Array(caseIndex, budget, total, startFactor)
        TestWriteReverseBinaryVector handle, seed: TestWriteReverseBinaryVector handle, values'''
TestReplace(old,new)
TestReplace('CDbl(meta(4)), CDbl(meta(6)), CDbl(meta(7)))','startFactor, CDbl(meta(6)), CDbl(meta(7)))')
TestReplace('maximumCalls = CLng((2# * UBound(zero) + 1#) * (CDbl(meta(8)) + 1#) + CDbl(meta(8)))',
    'maximumCalls = CLng((2# * UBound(zero) + 1#) * (CDbl(budget) + 1#) + CDbl(budget))')
TestReplace('solver.TrySolve(constrained, zero, maximumCalls, CLng(meta(8)), tolerance','solver.TrySolve(constrained, zero, maximumCalls, budget, tolerance')
TestReplace('factor = CDbl(meta(4)): value = 0#','factor = startFactor: value = 0#')
TestReplace('        mReverseSolverTest.EndNewtonLineTest',
'''        total = total + solver.Steps
        If total > CLng(meta(8)) Then Err.Raise 5, , "Материальная пара превысила общий бюджет."
        If caseIndex = 1 Then
            fullReady = (solved And materialReady And mappedReady And norm <= tolerance And Abs(value) <= 0.000000000002)
            If fullReady Then fullField = field: fullFactor = factor
        End If
        mReverseSolverTest.EndNewtonLineTest''')
TestReplace('    Close #handle\n    mReverseSolverTest.CopyAccepted checked',
    '    TestWriteReverseBinaryVector handle, Array(0#, total, mReverseSolverTest.AcceptedRevision)\n    Close #handle\n    mReverseSolverTest.CopyAccepted checked')
assert max(map(len,entry.splitlines()))<1024
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
ancestor=root/'MidpointPairCompilationAncestor';ancestor.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(ancestor/name)
updates={};prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n"
anchor=prefix*23+'\n\nPublic Function TestRunControlledSecantPrefix()';assert code.count(anchor)==1
updates['tests/modTestCoupledElasticState.bas']=code.replace(anchor,prefix*24+'\n\n'+entry+'\n\n\nPublic Function TestRunControlledSecantPrefix()',1).replace('\n','\r\n').encode('utf-8-sig')
updates['TestRunMaterialCoordinatePair.bas.txt']=(entry+'\n').replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
updates['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'modTestCoupledElasticState' -EntryPointNames 'TestRunMaterialCoordinatePair'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig')
old='modTestCoupledElasticState.TestSolveRecordedYieldCoordinate';assert runner.count(old)==1
updates[name]=runner.replace(old,'modTestCoupledElasticState.TestRunMaterialCoordinatePair',1).replace('\n','\r\n').encode('utf-8-sig')
updates.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='SequentialMaterialPairWithNativeMidpointSeedAndShared30',actual_import_this_pass=['modTestCoupledElasticState'],
    actual_component_count=257,all256_numerical_components_exact973=True,all_working_methods_unchanged=True,
    full_budget15_half_uses_true_remaining30=True,half_seed_native_midpoint_accepted_and_verified_full=True,
    half_lambda_prediction_midpoint_accepted_and_full=True,first_initial_residual_bitexact_actual968_guard_preserved=True,
    original_materials_force_path_tolerances_and_targets_exact973=True,no_material_history_acceptance=True,
    native_lambda_uncertainty_not_recorded=True,not_production_recovery_or_physical_capacity=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'NEW TEST ENTRY ONLY;NUMERICAL256 EXACT973;FULL15+ACTUAL REMAINING;NATIVE MIDPOINT HALF;NO HISTORY ACCEPTANCE')
