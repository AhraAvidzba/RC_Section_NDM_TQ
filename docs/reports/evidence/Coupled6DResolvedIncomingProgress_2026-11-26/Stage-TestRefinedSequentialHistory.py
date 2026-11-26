"""ДЛЯ ТЕСТОВ. Уточняет только равновесие второго материального подшага955, не меняя исходную приёмку."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold955SequentialEventHistory_20261126';root=base/'Cold956RefinedSequentialHistory_20261126'
proof=json.loads((parent/'SequentialEventHistoryProof.json').read_bytes())
assert proof['second_event_solved'] and proof['comparison']['same_terminal_yield_surface']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
ancestor=root/'RefinedSequentialCompilationAncestor';ancestor.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(ancestor/name)
entry=(here/'TestRunSequentialEventHistory.bas.txt').read_text('utf-8-sig')
entry=entry.replace('TestRunSequentialEventHistory','TestRunRefinedSequentialHistory')
old='CDbl(context(3)), CDbl(context(4)), CLng(context(5)) - total, eventField';assert entry.count(old)==1
entry=entry.replace(old,'0.0001 * CDbl(context(3)), 0.0001 * CDbl(context(4)), CLng(context(5)) - total, eventField',1)
entry=entry.replace('TOTAL_SEQUENTIAL_EVENT_HISTORY_TEST','TOTAL_REFINED_SEQUENTIAL_HISTORY_TEST')
updates={};name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n"
anchor=prefix*18+'\n\nPublic Function TestRunSequentialEventHistory()';assert code.count(anchor)==1
updates[name]=code.replace(anchor,prefix*19+'\n\n'+entry.rstrip()+'\n\n\nPublic Function TestRunSequentialEventHistory()',1).replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
updates['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'modTestCoupledElasticState' -EntryPointNames 'TestRunRefinedSequentialHistory'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig')
old='modTestCoupledElasticState.TestRunSequentialEventHistory';assert runner.count(old)==1
updates[name]=runner.replace(old,'modTestCoupledElasticState.TestRunRefinedSequentialHistory',1).replace('\n','\r\n').encode('utf-8-sig')
updates.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes(),
    'TestRunRefinedSequentialHistory.bas.txt':entry.replace('\n','\r\n').encode('utf-8-sig')})
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():
    if name.endswith('.bas'):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='RefineOnlySecondSequentialMaterialEventEquilibrium',actual_import_this_pass=['modTestCoupledElasticState'],
    actual_component_count=257,all_numerical_classes_exact955=True,second_event_tolerance_factor=.0001,
    half_and_full_source_unchanged=True,original_force_acceptance_and_total30_budget_preserved=True,
    geometric_guard_unchanged=True,temporary_history_restored_twice=True,no_production_history_acceptance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'ONLY SECOND EVENT SOLVE PRECISION1e-4;ORIGINAL ACCEPTANCE UNCHANGED')
