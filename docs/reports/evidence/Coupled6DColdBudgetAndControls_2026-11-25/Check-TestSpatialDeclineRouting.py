"""ДЛЯ ТЕСТОВ. Проверяет три маршрута возврата пространственной ветви без изменения критерия физического предела."""
from pathlib import Path
import json,re

base=Path('D:/CodexScratch/RC_Section_NDM_TQ');root=base/'Cold857SpatialDeclineRecovery_20261125';parent=base/'Cold855BalancedInitialControl_20261125'
path='src/Solver/CCoupledDeformationContinuation.cls';old=(parent/path).read_text('utf-8-sig');new=(root/path).read_text('utf-8-sig')
line="ElseIf reinforced And mSolver.UsesSpatialHistory Then ' Пространственный возврат спада применим и к обобщённому закручиванию, без подтверждения самого пика."
assert new.count(line)==1
restored=new.replace(line,'ElseIf reinforced And mSolver.UsesSpatialHistory And (mUsesTransverseControl Or mUsesFieldControl) Then',1)
line='If reinforced And (mSolver.UsesSpatialHistory Or mUsesTransverseControl Or mUsesFieldControl) Then'
assert restored.count(line)==2
restored=restored.replace(line,'If reinforced And (mUsesTransverseControl Or mUsesFieldControl) Then')
assert restored==old
methods=lambda code:dict((n,m) for m,n in re.findall(r'(?ms)^((?:Private|Public|Friend) (?:Function|Sub|Property Get) (\w+)\b.*?^End (?:Function|Sub|Property))',code))
a,b=methods(old),methods(new);assert set(a)==set(b)
assert {n for n in a if a[n]!=b[n]}=={'Trace'}
assert a['ReplayUnverifiedDecline']==b['ReplayUnverifiedDecline']
assert 'Private Const PROBE_LIMIT As Long = 96' in new
for file in (parent/'src').rglob('*'):
    if file.is_file() and file.relative_to(parent).as_posix()!=path:
        assert file.read_bytes()==(root/file.relative_to(parent)).read_bytes(),file
assert (parent/'tests/modTestCoupledElasticState.bas').read_bytes()==(root/'tests/modTestCoupledElasticState.bas').read_bytes()
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
assert audit['components']==256 and audit['imported_this_pass']==['CCoupledDeformationContinuation']
assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==255
selector=json.loads((parent/'BalancedInitialControlProof.json').read_bytes())
assert selector['only_Trace_initial_selection_call_changed'] and len(selector['cases'])==7
proof=dict(actual_components=256,one_controller_import_other255_exact=True,only_three_Trace_routing_conditions_changed=True,
    existing_ReplayUnverifiedDecline_exact=True,all_other_methods_Newton_arc_and_materials_exact855=True,
    original96_and30_limits=True,initial_balanced_selector_unchanged=True,
    old_transverse_and_field_paths_preserved=True,spatial_generalized_control_can_use_existing_decline_replay=True,
    no_new_peak_or_physical_limit_certificate=True,cold_workflow_checked_separately=True,physical_capacity_certified=False,whole_goal_completed=False)
(root/'SpatialDeclineRoutingProof.json').write_text(json.dumps(proof,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('EXACT THREE TRACE ROUTING CONDITIONS;256 ACTUAL;REPLAY/PHYSICS/BUDGETS UNCHANGED')
