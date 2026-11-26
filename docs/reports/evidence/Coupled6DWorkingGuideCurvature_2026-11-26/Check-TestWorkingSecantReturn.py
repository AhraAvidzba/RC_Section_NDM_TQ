"""ДЛЯ ТЕСТОВ. Независимо проверяет штатный возврат947, суммарный бюджет, исходные61 уравнение и отрицательную бюджетную проверку."""
from pathlib import Path
from runpy import run_path
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold947WorkingSecantReturn_20261126'
template=here/'Check-TestRecordedYieldCoordinate.py';code=template.read_text('utf-8-sig')
cut="r=R((root/'Native/RecordedYieldCoordinate.bin').read_bytes())";assert code.count(cut)==1;code=code[:code.index(cut)]
replacements={"root=base/'Cold936RecordedYieldCoordinate_20261126'":"root=base/'Cold947WorkingSecantReturn_20261126'",
    "audit['imported_this_pass']==['modTestCoupledElasticState']":"audit['imported_this_pass']==['CCoupledNewtonSolver','modTestCoupledElasticState']",
    "len(audit['other_actual_components_exact_parent_except_identifier_case'])==256":"len(audit['other_actual_components_exact_parent_except_identifier_case'])==255",
    "if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()":"if path.is_file() and path.name!='CCoupledNewtonSolver.cls':assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()"}
for before,after in replacements.items():assert code.count(before)==1;code=code.replace(before,after,1)
scope={'__file__':str(template),'__name__':'working_secant_independent_test'};exec(compile(code,str(template),'exec'),scope)
R=scope['R'];meta=scope['meta'];source=scope['source'];projection=scope['projection'];op=scope['op'];axis=scope['axis'];L=scope['L']
r=R((root/'Input/GuideReturnEffort.bin').read_bytes());r.position=0;cost=r.VectorTest();assert r.position==len(r.raw) and cost[0]==17
r=R((root/'Native/WorkingGuideReturn.bin').read_bytes());assert r.raw[:16]==b'NDMGUIDERET00001' and r.VectorTest().tolist()==[2.]
cases=[]
for index in (1,2):
    info,effort=r.VectorTest(),r.VectorTest();assert len(info)==10 and len(effort)==8 and info[0]==index
    assert info[3]==(cost[0] if index==1 else meta[7]) and info[6]==meta[7]
    assert info[3]<=effort[0]<=meta[7] and effort[2]>=info[3] and effort[3]>=cost[2] and effort[4]>=cost[1]
    text=(root/f'Native/WorkingGuideReturn.Case{index}.txt').read_text('cp1251')
    item=dict(case=index,native_solved=bool(info[1]),lambda_value=float(info[2]),guide_steps=int(info[3]),
        total_steps_including_guide=int(effort[0]),total_budget=int(meta[7]),effort=effort.tolist(),
        internal_status_snapshot=int(info[7]),result_code_snapshot=int(info[8]),diagnostics=text,
        independently_solves_original_control=False)
    if info[1]:
        field,actual,values,forces=[r.VectorTest() for _ in range(4)];assert len(field)==60 and len(actual)==6 and len(values)==61 and len(forces)==6
        stress=np.zeros((3,43,6));seen=set()
        for _ in range(129):
            point,side,error=r.VectorTest();sigma=r.VectorTest();key=int(point),int(side)
            assert error==0 and key not in seen and len(sigma)==6;seen.add(key);stress[int(side)+1,int(point)-1]=sigma
        a=scope['TestAssemble'](field,info[2],stress)
        phys=float(max(abs(values[:56]-a['residual'][:56])))
        gauge=float(max(abs(values[56:60]-meta[8]*a['residual'][56:])))
        force=float(max(abs(forces-a['six_force_residual_vector'])))
        control=float(meta[8]*(projection@field-meta[4]));cerr=float(abs(values[-1]-control))
        assert max(phys,gauge,force,cerr)<2e-11
        scales=np.array([1/L[0,0],1/L[1,1],1/L[1,1],1/L[1,1],1/L[0,0],1/L[0,0]])
        target=source['offset']+info[2]*source['direction'];native_force=(actual-target)/scales
        assert max(abs(native_force-forces))<2e-11
        bound=float(source['options'][3]+source['options'][4]*max(abs((L@target)[:56])))
        norm=float(max(a['physical_norm'],a['six_force_norm'],meta[8]*max(abs(a['residual'][56:])),abs(control)))
        strain=np.einsum('spcd,d->spc',op,field);bars=np.einsum('spc,pc->sp',strain[:,25:],axis[25:]);long=axis[25:,2]==1
        steel=float(np.max(abs(bars)/np.where(long[None,:],np.where(bars>=0,.025,.018),.02)))
        verified=bool(norm<=bound+1e-12 and steel<=1)
        assert verified and 'полный/половинный возврат' in text and 'прогноз исходной цели по секущей' in text
        item.update(independently_solves_original_control=verified,full_original_residual_MPa=norm,
            original_final_bound_MPa=bound,original_control_residual_MPa=control,six_force_residual_MPa=a['six_force_norm'],
            steel_strain_utilization=steel,all_material_returns_available=True,
            native_full_half_and_lambda_uncertainty_guards_passed=True,half_state_not_independently_exported=True)
    if index==2:
        assert not info[1] and effort[0]==meta[7] and effort[5]==0 and effort[3]==cost[2] and effort[4]==cost[1]
        item.update(synthetic_exhausted_budget=True,no_new_iterations_or_hybrid_matrices=True)
    cases.append(item)
assert r.position==len(r.raw)
proof=dict(native=scope['native'],actual_components=257,declared_components_changed=2,other255_components_preserved=True,
    source_history25_and_all_materials_unchanged=True,initial_guide938_residual_bitexact_guard=True,
    working_original_return_used=True,guide17_steps_in_total30=True,earlier_primary_failure_effort_not_included=True,
    no_half_uncertainty_direction_or_steel_guard_relaxed=True,no_material_history_acceptance=True,cases=cases,
    verified_working_original_returns=sum(c['independently_solves_original_control'] for c in cases),
    physical_capacity_certified=False,cold_gain_claimed=False,whole_goal_completed=False)
(root/'WorkingSecantReturnProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
for c in cases:print('CASE',c['case'],'NATIVE',c['native_solved'],'VERIFIED',c['independently_solves_original_control'],
    'LAMBDA',c['lambda_value'],'TOTAL STEPS',c['total_steps_including_guide'],'R',c.get('full_original_residual_MPa'))
