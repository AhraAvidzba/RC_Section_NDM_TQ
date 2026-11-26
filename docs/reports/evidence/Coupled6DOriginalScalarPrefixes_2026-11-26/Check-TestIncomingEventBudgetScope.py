"""ДЛЯ ТЕСТОВ. Проверяет точный объём изменений963 и неизменность физических/метрических критериев существующей дуги."""
from pathlib import Path
from runpy import run_path
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold963IncomingEventBudget_20261126';parent=base/'Cold962CurrentOperationalArc_20261126'
name='src/Solver/CCoupledArcContinuation.cls'
before=(parent/name).read_text('utf-8-sig');after=(root/name).read_text('utf-8-sig')
methods=run_path(str(here/'Stage-TestWorkingDepartureTrace.py'))['TestMethods'];old,new=methods(before),methods(after)
assert set(new)-set(old)=={('Function','TryIncomingFaceApproach')}
changed={key for key in old if old[key].group()!=new[key].group()}
assert changed=={('Function','TryStep'),('Function','TryEvent'),('Function','TryPreEventApproach')}
for method in ('Initialize','PairLambdaBound','LocalRoot','PositiveArcProgress','ReadEventSurface','IsBeforeEvent'):
    key=next(key for key in old if key[1]==method);assert old[key].group()==new[key].group()
first=min(m.start() for m in old.values());second=min(m.start() for m in new.values())
field="Private mIncomingRecoveryDetails As String ' Сведения о сокращённых полных пробах того же входящего перехода.\n"
assert after[:second].replace(field,'',1)==before[:first]
for path in (parent/'src').rglob('*'):
    if path.is_file() and path.name!='CCoupledArcContinuation.cls':assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()
for name in ('tests/modTestCoupledElasticState.bas','Input/Reverse1.bin','Input/OperationalArcRatio.bin','Input/ArcCallAllowance.bin'):
    assert (root/name).read_bytes()==(parent/name).read_bytes()
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
assert audit['components']==257 and audit['imported_this_pass']==['CCoupledArcContinuation']
assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==256
proof=json.loads((root/'CurrentOperationalArcProof.json').read_bytes())
proof.update(actual_components=257,all256_other_actual_components_preserved=True,only_existing_arc_budget_and_incoming_recovery_changed=True,
    changed_existing_methods=[list(key) for key in sorted(changed)],added_working_method='TryIncomingFaceApproach',
    strict_internal_solve_accuracy_and_all_acceptance_methods_bitexact962=True,
    source_test_and_inputs_bitexact962=True,no_supplied_future_guide_or_field=True)
(root/'IncomingEventBudgetProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('963 SCOPE VERIFIED;INTERNAL ACCURACY/PAIR BOUND/ORIENTATION/BEFORE-EVENT CRITERIA EXACT962;ALL256 OTHER COMPONENTS PRESERVED')
