"""ДЛЯ ТЕСТОВ. Независимо проверяет выбранное общим поиском событие952, исходные силовые уравнения и материальную границу."""
from pathlib import Path
from runpy import run_path
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold952DetectedLoadingBoundary_20261126';parent=base/'Cold949WorkingGuidePair_20261126'
template=here/'Check-TestRecordedYieldCoordinate.py';code=template.read_text('utf-8-sig')
cut="r=R((root/'Native/RecordedYieldCoordinate.bin').read_bytes())";assert code.count(cut)==1;code=code[:code.index(cut)]
replacements={"root=base/'Cold936RecordedYieldCoordinate_20261126'":"root=base/'Cold952DetectedLoadingBoundary_20261126'",
    "if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()":"if path.is_file() and path.name!='CCoupledNewtonSolver.cls':assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()"}
for before,after in replacements.items():assert code.count(before)==1;code=code.replace(before,after,1)
scope={'__file__':str(template),'__name__':'detected_loading_boundary_test'};exec(compile(code,str(template),'exec'),scope)
for path in (parent/'src').rglob('*'):
    if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()
R=scope['R'];meta=scope['meta'];source=scope['source'];L=scope['L'];op=scope['op'];axis=scope['axis']
r=R((root/'Native/DetectedLoadingBoundary.bin').read_bytes());assert r.raw[:16]==b'NDMDETECTEV00001'
info,event,effort=r.VectorTest(),r.VectorTest(),r.VectorTest();count=int(r.VectorTest()[0]);capture=[r.VectorTest() for _ in range(count)]
assert len(info)==3 and len(effort)==8 and effort[0]<=source['options'][2]
proof=dict(native=scope['native'],actual_components=257,all256_other_components_preserved=True,
    numerical_classes_exact949=True,selected_by_existing_generic_detector=True,no_point_passed_to_API=True,
    native_solved=bool(info[0]),lambda_value=float(info[1]),selected_event=event.tolist(),effort=effort.tolist(),
    native_capture=[c.tolist() for c in capture],independently_equilibrated=False,
    source_history25_unchanged=True,no_material_history_acceptance=True,physical_capacity_certified=False)
if info[0]:
    field,values,forces=r.VectorTest(),r.VectorTest(),r.VectorTest();stress=np.zeros((3,43,6));seen=set()
    for _ in range(129):
        point,side,error=r.VectorTest();sigma=r.VectorTest();key=int(point),int(side)
        assert error==0 and key not in seen;seen.add(key);stress[int(side)+1,int(point)-1]=sigma
    a=scope['TestAssemble'](field,info[1],stress)
    physical=float(max(abs(values[:56]-a['residual'][:56])));gauge=float(max(abs(values[56:60]-meta[8]*a['residual'][56:])))
    force=float(max(abs(forces-a['six_force_residual_vector'])));assert max(physical,gauge,force)<2e-11
    goal=source['offset']+info[1]*source['direction'];bound=float(source['options'][3]+source['options'][4]*max(abs((L@goal)[:56])))
    norm=float(max(a['physical_norm'],a['six_force_norm'],meta[8]*max(abs(a['residual'][56:]))))
    assert norm<=bound+2e-11 and abs(values[-1])<2e-11
    assert len(event)==7 and int(event[1])==18 and int(event[2])==0
    material=float(scope['TestYield'](int(event[1]),int(event[2]),field));assert abs(material)<2e-10
    strain=np.einsum('spcd,d->spc',op,field);bars=np.einsum('spc,pc->sp',strain[:,25:],axis[25:]);long=axis[25:,2]==1
    steel=float(np.max(abs(bars)/np.where(long[None,:],np.where(bars>=0,.025,.018),.02)));assert steel<=1
    proof.update(independently_equilibrated=True,independent_total_residual_MPa=norm,final_bound_MPa=bound,
        independent_event_yield_residual_MPa=material,maximum_steel_strain_utilization=steel,
        six_force_residual_MPa=a['six_force_norm'],field=field.tolist(),values=values.tolist(),
        all_local_material_returns_available=True,event_before_known_half949=bool(info[1]<1.1038129802817196))
assert r.position==len(r.raw)
(root/'DetectedLoadingBoundaryProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('GENERIC EVENT',event.tolist(),'NATIVE',proof['native_solved'],'VERIFIED',proof['independently_equilibrated'],
    'LAMBDA',proof['lambda_value'],'STEPS',int(effort[0]),'R',proof.get('independent_total_residual_MPa'),'F',proof.get('independent_event_yield_residual_MPa'))
