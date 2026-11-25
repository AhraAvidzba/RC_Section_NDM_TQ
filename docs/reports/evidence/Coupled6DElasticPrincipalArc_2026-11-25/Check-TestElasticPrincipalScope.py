"""ДЛЯ ТЕСТОВ. Сопоставляет все тела методов кандидатов с точными родительскими источниками без нормализации арифметики."""
from pathlib import Path
import json,re

base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
old=base/'Cold834UnresolvedArcScan_20261125';point=base/'Cold837ElasticPrincipalEvent_20261125';arc=base/'Cold838ElasticPrincipalArc_20261125'
def TestMethods(path):
    """Выделяет тела методов, сохраняя регистр, строки, комментарии, условия и арифметику."""
    code=path.read_text('utf-8-sig')
    rows=re.findall(r'(?ms)^((?:Private|Friend|Public) (?:Function|Sub|Property Get) (\w+)\b.*?^End (?:Sub|Function|Property))',code)
    assert len(rows)==len({name for _,name in rows})
    return {name:body for body,name in rows}
results=[]
for before,after,name,changed,added in [
    (old,point,'src/Materials/CCoupledConcreteResponse.cls',set(),{'TryReadElasticPrincipalBoundary'}),
    (old,point,'src/Materials/CCoupledReinforcedResponse.cls',set(),{'TryReadElasticPrincipalBoundary'}),
    (old,point,'src/Solver/CCoupledNewtonSolver.cls',{'CreateMaterialConstraint','ReadTransitionValue','FollowYieldEventCoordinate'},
        {'TryElasticPrincipalBoundary','TryReadElasticPrincipalSurface'}),
    (point,arc,'src/Solver/CCoupledNewtonSolver.cls',{'TryLoadingYieldBoundary','TryElasticPrincipalBoundary'}, {'SelectElasticPrincipalStation'}),
    (point,arc,'src/Solver/CCoupledArcContinuation.cls',{'TryEvent','TryPreEventApproach'}, {'ReadEventSurface','IsBeforeEvent'})]:
    a,b=TestMethods(before/name),TestMethods(after/name)
    assert set(b)==set(a)|added,(name,set(b)-set(a))
    actual={n for n in a if a[n]!=b[n]};assert actual==changed,(name,actual)
    results.append(dict(source=before.name,target=after.name,file=name,changed=sorted(changed),added=sorted(added),
        unchanged_methods=len(a)-len(changed)))
name='src/Solver/CCoupledNewtonSolver.cls'
a,b=TestMethods(old/name),TestMethods(point/name)
new=b['FollowYieldEventCoordinate'].replace('station.PrincipalIndex, target, station.ElasticPrincipal)','0, target)')
assert new==a['FollowYieldEventCoordinate'],'Full material-event solve math must be unchanged'
name='src/Solver/CCoupledArcContinuation.cls';a,b=TestMethods(point/name),TestMethods(arc/name)
new=b['TryEvent'].replace('ReadEventSurface(eventInfo, half, halfYield)',
    'mSolver.TryReadLoadingSurface(CLng(eventInfo(1)), CLng(eventInfo(2)), half, halfYield)')
new=new.replace('IsBeforeEvent(eventInfo, halfYield)','halfYield < 0#')
assert new==a['TryEvent'],'Event half-pair arithmetic/acceptance budgets changed'
new=b['TryPreEventApproach']
for field,value in [('candidate','candidateYield'),('probe','probeYield')]:
    new=new.replace(f'ReadEventSurface(eventInfo, {field}, {value})',
        f'mSolver.TryReadLoadingSurface(CLng(eventInfo(1)), CLng(eventInfo(2)), {field}, {value})')
    new=new.replace(f'IsBeforeEvent(eventInfo, {value})',f'{value} < 0#')
assert new==a['TryPreEventApproach'],'Pre-event half-pair arithmetic/acceptance budgets changed'
for name in ['src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledReinforcedResponse.cls',
    'src/Materials/CCoupledConcreteReturn.cls','src/Solver/CCoupledMaterialConstraint.cls',
    'src/Solver/CCoupledDeformationContinuation.cls','Input/Reverse1.bin']:
    assert (point/name).read_bytes()==(arc/name).read_bytes(),name
proof=dict(comparisons=results,all_original_material_evaluation_methods_unchanged=True,
    material_event_projected_solve_math_unchanged=True,arc_half_pair_and_pre_event_acceptance_math_unchanged=True,
    main_TryStep_arc_plane_and_exact_restore_unchanged=True,
    controller_physical_checks_unchanged=True,physical_capacity_certified=False)
(arc/'ElasticPrincipalScopeProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('METHOD SCOPE VERIFIED:',results)
