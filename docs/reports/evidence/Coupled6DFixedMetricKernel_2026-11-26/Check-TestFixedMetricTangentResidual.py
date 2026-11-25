"""ДЛЯ ТЕСТОВ. Собирает физическую невязку фактического неудачного касательного прогноза по независимо прочитанным напряжениям."""
from pathlib import Path
from runpy import run_path
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold869FixedMetricNextScan_20261126';parent=base/'Cold867FixedMetricKernel_20261126'
scan=json.loads((root/'ColdAcceptedStressContinuityProof.json').read_bytes());assert scan['postcommit_stress_drift_within1e_minus11_MPa']
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((parent/'Native/LoadingWalk.Guide003.bin').read_bytes());r.VectorTest()
for size in (3,3):
    if r.VectorTest()[0]:
        for _ in range(size):r.VectorTest()
assert r.VectorTest()[0]==1
meta,probe,tangent,movement=[r.VectorTest() for _ in range(4)];assert r.position==len(r.raw)
warm=json.loads((parent/'FixedMetricKernelProof.json').read_bytes());ratio=warm['attempts'][-1]['ratio']
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)]);rows=run_path(str(here/'Validate-TestHardPrepared.py'))['TestRows'](root/'Native/ArcScan.csv')
levels=(0.,.03125,.0625,.125,.25,.5,.75,1.);stress={level:np.zeros((3,43,6)) for level in levels};seen=set()
for row in rows:
    if row[0]!='S' or int(row[1])!=3:continue
    _,case,level,point,side,failure,*values=row;level=float(level);point=int(point)-1;side=int(side)+1
    assert int(failure)==0 and (level,point,side) not in seen;seen.add((level,point,side));stress[level][side,point]=np.array(values,float)
assert len(seen)==8*129
accepted=source['fields'][-1];reference=source['lambdas'][-1];E=source['calibration'][0];results=[]
for level in levels:
    field=accepted+level*ratio*movement;factor=reference+level*ratio*meta[1]
    strain=np.einsum('spcd,d->spc',op,field);bar=200000*np.einsum('spc,pc->sp',strain[:,25:],axis[25:])
    assert np.max(abs(bar))<240
    net=stress[level].copy();net[:,25:]=bar[:,:,None]*axis[None,25:]-net[:,25:]
    target=source['offset']+factor*source['direction']
    residual=g@field+np.einsum('pcd,pc->d',w,net[1])+np.einsum('pcd,pc->d',v,(net[2]-net[0])/.002)-L@target
    norm=float(max(max(abs(residual[:56])),E*max(abs(residual[56:]))));dominant=int(np.argmax(abs(residual[:56])))+1
    results.append(dict(level=level,lambda_value=float(factor),residual_MPa=norm,dominant_physical_row=dominant,residual_vector=residual.tolist()))
base_res=np.array(results[0]['residual_vector']);delta=[]
for item in results[1:]:
    dr=np.array(item['residual_vector'])-base_res;dn=float(max(max(abs(dr[:56])),E*max(abs(dr[56:]))))
    item['incremental_residual_MPa']=dn;item['incremental_residual_divided_by_level_MPa']=dn/item['level'];delta.append(dn)
proof=dict(source20_true=True,same_actual_failed_outgoing_tangent=True,no_new_material_or_equilibrium_acceptance=True,
    last_failed_ratio=ratio,tangent_signed_lambda_increment=float(meta[1]),scan=results,
    residual_ratio_second_to_first_positive_level=delta[1]/delta[0],
    equilibrium_residual_not_a_failure_or_limit_criterion=True,physical_capacity_certified=False)
(root/'FailedTangentResidualProof.json').write_text(json.dumps(proof,indent=2,sort_keys=True)+'\n',encoding='utf-8')
for row in results:print(row['level'],row['residual_MPa'],row['dominant_physical_row'],row.get('incremental_residual_divided_by_level_MPa'))
print('SECOND/FIRST RESIDUAL RATIO',proof['residual_ratio_second_to_first_positive_level'])
