"""ДЛЯ ТЕСТОВ. Проверяет новые промежуточные шаги по разрешимому полю, сохранённые силовые критерии и отрицательные случаи."""
from pathlib import Path
import json,numpy as np
here=Path(__file__).resolve().parent;script=here/'Check-TestBoundedIncomingSteps.py';code=script.read_text('utf-8-sig')
old="noise=float(8*context[1]*(1+abs(previous_lambda)));assert tag[2]>previous_lambda+noise+tag[6]"
new="""noise=float(8*context[1]*(1+abs(previous_lambda)))
            load_increase_resolved=bool(tag[2]>previous_lambda+noise+tag[6])
            advance=float(p@(field-accepted));guide_advance=float(p@(guide-accepted))
            tolerance=cycle['attempts'][-1]['pairs'][-1]['full']['bound_MPa'];resolution=16*tolerance/E
            assert abs(advance)>resolution and advance*guide_advance>0
            if not load_increase_resolved:
                assert abs(advance)<=abs(guide_advance)+resolution+1e-15
                assert TestNorm(field-accepted)<=TestNorm(guide-accepted)+resolution+1e-15"""
changes={
 "root=base/'Cold958BoundedIncomingSteps_20261126';parent=base/'Cold957FourPartEventHistory_20261126'":"root=base/'Cold960ResolvedFieldProgress_20261126';parent=base/'Cold958BoundedIncomingSteps_20261126'",
 "audit['imported_this_pass']==['modTestCoupledElasticState']":"audit['imported_this_pass']==['CCoupledNewtonSolver','modTestCoupledElasticState']",
 "len(audit['other_actual_components_exact_parent_except_identifier_case'])==256":"len(audit['other_actual_components_exact_parent_except_identifier_case'])==255",
 "if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()":"if path.is_file() and path.name!='CCoupledNewtonSolver.cls':assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()",
 old:new,
 "source_force_and_half_and_steel_checks_passed=True)":"source_force_and_half_and_steel_checks_passed=True,load_increase_resolved=load_increase_resolved,field_advance=advance,field_resolution=resolution)",
}
for old,new in changes.items():assert code.count(old)==1;code=code.replace(old,new,1)
env={'__file__':str(script),'__name__':'resolved_field_progress_test'};exec(compile(code,str(script),'exec'),env)
root=env['root'];parent=env['parent'];proof=json.loads((root/'BoundedIncomingStepsProof.json').read_bytes())
before=(parent/'src/Solver/CCoupledNewtonSolver.cls').read_text('utf-8-sig')
after=(root/'src/Solver/CCoupledNewtonSolver.cls').read_text('utf-8-sig')
block=(root/'FieldProgressCriteriaBlock.txt').read_text('utf-8-sig');assert after.count(block)==1
original='    If returnedLambda <= referenceLambda + 8# * relativeTolerance * (1# + Abs(referenceLambda)) + uncertainty Then GoTo RejectedReturn'
assert after.replace(block,original,1)==before
r=env['R']((root/'Native/IncomingProgressGuards.bin').read_bytes());assert r.raw[:16]==b'NDMINCGUARDS0001'
guards=[]
for index in range(1,4):
    info,effort=r.VectorTest(),r.VectorTest();assert info[0]==index and not info[1] and effort[0]<=env['context'][2]
    if index==3:assert effort[0]==env['context'][2]
    guards.append(dict(case=index,rejected=True,steps=int(effort[0])))
assert r.position==len(r.raw)
proof.update(all_numerical_sources_exact957=False,all256_other_components_preserved=False,all255_other_components_preserved=True,
    original_force_steel_shape_and_uncertainty_guards_preserved=False,
    original_force_steel_shape_positive_lambda_and_uncertainty_calculation_preserved=True,
    resolved_load_increase_requirement_replaced_for_intermediate_field_continuation=True,
    resolved_directed_field_and_local_guide_radius_independently_verified=True,
    negative_zero_reverse_and_exhausted_budget_guards=guards,
    accepted_steps_with_unresolved_load_increase=sum(not c['accepted']['load_increase_resolved'] for c in proof['cycles'] if 'accepted' in c))
proof.pop('all256_other_components_preserved')
proof.pop('original_force_steel_shape_and_uncertainty_guards_preserved')
for name in ('BoundedIncomingStepsProof.json','ResolvedFieldProgressProof.json'):
    (root/name).write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('FIELD PROGRESS ACCEPTED',proof['accepted_steps'],'INCLUDING UNRESOLVED LOAD INCREASE',proof['accepted_steps_with_unresolved_load_increase'],'NEGATIVE GUARDS',len(guards))
