"""ДЛЯ ТЕСТОВ. Проверяет холодный автоматический поиск с существующей собранной касательной и проверенным внутренним остановом."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold732ProvidedJacobianCapacity_20261124';parent=base/'Cold723PriorBranchCapacity_20261124'
candidate=base/'Cold731OuterToleranceShoulder_20261124'
assert json.loads((candidate/'ShoulderWitnessProof.json').read_bytes())['quality_confirmed']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ProvidedCapacityCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
values={};prep=json.loads((root/'Preparation.json').read_bytes())
name='src/Solver/CCoupledNewtonSolver.cls';code=(candidate/name).read_text('utf-8-sig')
code,count=re.subn(r'(?m)^\s*modTestCoupledElasticState\.TestRecordReplayPhase[^\n]*\n','',code)
assert count==7
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='RunTestCoupledPostCrackCapacity = RunTestCoupledPostCrackTension(evidenceRoot, True)';assert code.count(anchor)==1
code=code.replace(anchor,'RunTestCoupledPostCrackCapacity = RunTestCoupledPostCrackTension(evidenceRoot, True, False, 1)',1)
method,=re.findall(r'(?ms)^Public Function RunTestCoupledPostCrackTension\(.*?^End Function',code)
changed=method
start=changed.index('    solved = branch.CopyPeak(peak, accuracy, design)\n')
finish=changed.index('\nDone:\n',start)
checks=changed[start:finish];changed=changed[:start]+changed[finish:]
anchor='    report = report & "POSTCRACK_BRANCH: " & branch.ResultMeta.DiagnosticDetails & vbCrLf\n';assert changed.count(anchor)==1
changed=changed.replace(anchor,anchor+checks+'''
    If capacityOnly And Len(evidenceRoot) > 0 Then
        Dim outcomeHandleTest As Integer: outcomeHandleTest = FreeFile
        Open evidenceRoot & "\\Native\\AutomaticCapacityOutcome.txt" For Output As #outcomeHandleTest
        Print #outcomeHandleTest, report
        Print #outcomeHandleTest, "AUTOMATIC_RESULT: hasPeak=" & CStr(solved) & "; peak=" & TestOnsetNumber(peak) & _
            "; accuracy=" & TestOnsetNumber(accuracy) & "; designVerified=" & CStr(design) & "; steelEvent=" & CStr(branch.SteelEventPoint)
        Print #outcomeHandleTest, "PREEXPORT_TESTS: passed=" & CStr(passed) & "; failed=" & CStr(failed)
        Close #outcomeHandleTest
    End If
''',1)
anchor='    mStockProbeRootTest = vbNullString: Set mStockProbeObserverTest = Nothing\n';assert changed.count(anchor)==1
changed=changed.replace(anchor,'    If capacityOnly Then TestCaptureStockContinuation branch, branch.ResultMeta.DiagnosticDetails\n'+anchor,1)
assert checks in changed
code=code.replace(method,changed,1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames @('CCoupledNewtonSolver','modTestCoupledElasticState')")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='ColdCapacityWithProvidedJacobianAndInnerStop',actual_import_this_pass=['CCoupledNewtonSolver','modTestCoupledElasticState'],
    solver731_exact_except_readonly_phase_logging_removed=True,material722_and_controller719_exact=True,
    existing_test_tangent_mode=1,original_two_elastic_initial_fields_only=True,
    no_future_prescribed_fields_or_known_capacity_seed=True,original96_probes_30_iterations_and_accuracy_unchanged=True,
    numerical_checks_and_final_raw_history_flushed_before_tensor_export=True,original_capacity_assertions_exact=True,
    full_cold_from_zero=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('732 ACTUAL COLD AUTOMATIC CAPACITY; EXACT PHYSICS AND96/30; NO FUTURE FIELD SEED')
