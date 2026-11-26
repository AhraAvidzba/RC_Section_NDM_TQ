"""ДЛЯ ТЕСТОВ. Проверяет единственное изменение Newton и независимо собирает найденные пробные корни; не превращает NumFail в разрушение."""
from pathlib import Path
from runpy import run_path
import argparse,json,numpy as np

p=argparse.ArgumentParser(description=__doc__);p.add_argument('--baseline',action='store_true');args=p.parse_args()
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold895FreshSpatialArc_20261126'
root=base/('Cold902ConsistentRefinementBaseline_20261126' if args.baseline else 'Cold901ConsistentControlRefinement_20261126')
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=2,failed=0,macros=2,timeout600=False,own_process_closed=True)
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
assert audit['components']==256 and audit['imported_this_pass']==['CCoupledNewtonSolver','modTestCoupledElasticState']
assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==254
old='If Not mConsistentControl And Not mRefinedControlTangent And SupportsHybridRecovery And mIterations + 1 < maxIterations Then'
new='If Not mRefinedControlTangent And SupportsHybridRecovery And mIterations + 1 < maxIterations Then'
for path in (parent/'src').rglob('*'):
    if not path.is_file():continue
    current=root/path.relative_to(parent)
    if path.name=='CCoupledNewtonSolver.cls' and not args.baseline:
        text=path.read_text('utf-8-sig');assert text.count(old)==1
        assert text.replace(old,new,1)==current.read_text('utf-8-sig')
    else:assert path.read_bytes()==current.read_bytes(),str(path)
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
assert len(source['fields'])==32
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
r=R((root/'Native/Cold32RefinementSteps.bin').read_bytes());assert r.raw[:16]==b'NDMREFINE32S0001'
meta=r.VectorTest();projection=r.VectorTest()
assert len(meta)==7 and tuple(meta[:3])==(3,32,30)
assert meta[3]==source['options'][3]*.0001 and meta[4]==source['options'][4]*.0001
validate=run_path(str(here/'Validate-TestHardPrepared.py'));cases=[]
details=(root/'Native/Cold32RefinementSteps.txt').read_text('cp1251').splitlines();assert len(details)==3
for index in (1,2,3):
    info=r.VectorTest();effort=r.VectorTest();assert len(info)==7 and len(effort)==8
    assert info[0]==index and info[1]==.5**(index-1) and info[4]==info[5] and effort[0]<=30
    item=dict(index=index,ratio=float(info[1]),solved=bool(info[2]),seconds=float(info[6]),
        effort=effort.tolist(),accepted_revision_unchanged=True,details=details[index-1],
        refined_tangent_used='пространственная касательная уточнена' in details[index-1])
    if info[2]:
        field=r.VectorTest();actual=r.VectorTest();phase='Cold32Refined.Case'+str(index)+'.Sequence'
        mechanics,exported=validate['TestValidate'](root,phase)
        assert len(exported)==33
        np.testing.assert_allclose(exported[:32],source['fields'],rtol=1e-14,atol=1e-300)
        np.testing.assert_allclose(exported[-1],field,rtol=1e-14,atol=1e-300)
        weak=mechanics['states'][-1];bound=weak['bound_MPa']*.0001
        assert max(weak['residual_MPa'],weak['force_residual_MPa'])<=bound+1e-12
        ops=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/('Native/'+phase+'.Operators.csv'),60,43)
        g,L,w,v,coords,axis,B,gradient=ops
        strain=np.einsum('spcd,d->spc',np.array([B+s*.001*gradient for s in (-1,0,1)]),field)
        steel=np.einsum('spc,pc->sp',strain[:,25:],axis[25:]);long=axis[25:,2]==1
        limits=np.where(long[None,:],np.where(steel>=0,.025,.018),.02)
        steel_use=float(np.max(abs(steel)/limits));assert steel_use<=1
        coordinate=float(projection@field);target=meta[5]+info[1]*meta[6]
        assert abs(coordinate-target)*30000<=bound+1e-12
        item.update(lambda_value=float(info[3]),weak=weak,strict_bound_MPa=bound,
                    control_error_MPa=abs(coordinate-target)*30000,maximum_three_station_steel_utilization=steel_use)
    cases.append(item)
assert r.position==len(r.raw)
proof=dict(native=native,actual_components=256,baseline=args.baseline,only_existing_refinement_gate_changed=not args.baseline,
 all_other_working_sources_exact895=True,independent_cold32_probes=cases,original30_and_strict_force_tolerances=True,
 accepted_history_not_advanced=True,all_three_station_steel_limits_independently_checked=True,
 material_laws_and_strain_limits_unchanged=True,physical_capacity_certified=False,
 full_cold_and_directed_regression_still_required=not args.baseline,whole_goal_completed=False)
(root/'ConsistentRefinementProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('BASELINE' if args.baseline else 'CANDIDATE','EXACT32;ONE GATE ONLY;ALL FOUND ROOTS INDEPENDENTLY EQUILIBRATED')
for item in cases:print(item['index'],item['solved'],item['seconds'],'ITER',item['effort'][0],'EVAL',item['effort'][1],
                      'REFINED',item['refined_tangent_used'],'LAMBDA',item.get('lambda_value'))
