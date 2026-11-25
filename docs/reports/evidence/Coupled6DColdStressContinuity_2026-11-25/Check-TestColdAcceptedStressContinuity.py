"""ДЛЯ ТЕСТОВ. Сопоставляет напряжения до/после принятия того же поля и независимо собирает их влияние на равновесие."""
from pathlib import Path
from runpy import run_path
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold862ColdAcceptedStressContinuity_20261125';compiled=base/'Cold858BalancedInitialStressReplay_20261125';verified=base/'Cold859SpatialDeclineStressReplay_20261125'
native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
assert 'READONLY_PRIVATE_SCAN_COMPLETED: no new equilibrium search' in text
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(compiled/'docs/regression/Diagnostic.xlsm').read_bytes()
for suffix in ('.csv','.Operators.csv'):
    name='Native/EnvelopeEnergy.Accepted'+suffix;assert (root/name).read_bytes()==(verified/name).read_bytes(),name
validate=run_path(str(here/'Validate-TestHardPrepared.py'));mechanics,fields=validate['TestValidate'](root,'EnvelopeEnergy.Accepted')
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin');assert len(source['fields'])==19
np.testing.assert_allclose(fields,source['fields'],rtol=1e-14,atol=1e-300)
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
r=Reader((root/'Input/ArcScanFields.bin').read_bytes());r.position=0;assert r.VectorTest()[0]==3
predictions=[r.VectorTest() for _ in range(3)];assert r.position==len(r.raw)
np.testing.assert_array_equal(predictions[0],source['fields'][-1])
before=np.zeros((3,43,6));bars=np.zeros((3,43));seen=set()
for row in validate['TestRows'](root/'Native/EnvelopeEnergy.Accepted.csv'):
    if row[0]=='S' and int(row[1])==19:
        key=(int(row[2])+1,int(row[3])-1);assert key not in seen;seen.add(key)
        before[key]=np.array(row[4:10],float);bars[key]=float(row[12])
assert len(seen)==129
levels=(0.,.03125,.0625,.125,.25,.5,.75,1.);stresses={};yields={}
rows=validate['TestRows'](root/'Native/ArcScan.csv');assert len(rows)==3*8*129*2
for row in rows:
    kind,case,level,point,side=row[:5];key=(int(case),float(level),int(point),int(side));assert key[1] in levels
    if kind=='S':
        assert int(row[5])==0 and key not in stresses;stresses[key]=np.array(row[6:12],float)
    elif kind=='Y':
        assert row[5]=='True' and key not in yields;yields[key]=float(row[6])
    else:raise AssertionError(kind)
after=np.zeros_like(before)
for point in range(1,44):
    for side in (-1,0,1):
        zero=stresses[1,0.,point,side];after[side+1,point-1]=zero
        for level in levels:np.testing.assert_array_equal(stresses[1,level,point,side],zero)
        for case in (2,3):np.testing.assert_array_equal(stresses[case,0.,point,side],zero)
assert np.max(abs(bars[:,25:]))<240
# Все стальные ветви здесь упругие: то же точное поле даёт те же Es*epsilon без изменения пластической истории стали.
def TestEquilibrium(stress):
    """Вычитает бетон замещения арматуры и собирает исходные слабые уравнения и шесть интегральных усилий."""
    net=stress.copy();net[:,25:]=bars[:,25:,None]*axis[None,25:]-stress[:,25:]
    field=source['fields'][-1];target=source['fractions'][-1]*source['offset']+source['lambdas'][-1]*source['direction']
    residual=g@field+np.einsum('pcd,pc->d',w,net[1])+np.einsum('pcd,pc->d',v,(net[2]-net[0])/.002)-L@target
    area=coords[:,2];center=np.average(coords[:25,:2],axis=0,weights=area[:25]);x,y=(coords[:,:2]-center).T
    sigma=net[1];actual=np.array([area@sigma[:,2],(area*y)@sigma[:,2],(area*x)@sigma[:,2],area@(x*sigma[:,5]-y*sigma[:,4]),area@sigma[:,4],area@sigma[:,5]])
    scales=np.array([1/L[0,0],1/L[1,1],1/L[1,1],1/L[1,1],1/L[0,0],1/L[0,0]])
    norm=max(max(abs(residual[:56])),source['calibration'][0]*max(abs(residual[56:])))
    return dict(residual_MPa=float(norm),force_residual_MPa=float(max(abs((actual-target)/scales)))),residual,actual
old,oldr,oldf=TestEquilibrium(before);new,newr,newf=TestEquilibrium(after)
assert abs(old['residual_MPa']-mechanics['states'][-1]['residual_MPa'])<1e-11
crossings=[]
for case in (2,3):
    for point in range(1,44):
        for side in (-1,0,1):
            values=np.array([yields[case,l,point,side] for l in levels])
            if values[0]<-8e-11 and max(values)>8e-11:
                crossings.append(dict(case=case,point=point,side=side,yield_MPa=values.tolist()))
diff=after-before;largest=np.unravel_index(np.argmax(abs(diff)),diff.shape)
proof=dict(native=native,immutable_compiled_book_reused=True,actual_components=256,no_VBA_edit_or_new_compile=True,
 exact19_accepted_stresses_and_operators_equal_verified859=True,readonly_private_scan_completed=True,
 zero_increment_repeated24_times_per_point_is_bit_identical=True,
 maximum_postcommit_concrete_stress_difference_MPa=float(np.max(abs(diff))),largest_difference_side_point_component=[int(largest[0])-1,int(largest[1])+1,int(largest[2])+1],
 postcommit_stress_drift_within1e_minus11_MPa=bool(np.max(abs(diff))<=1e-11),
 before_commit_equilibrium=old,after_commit_zero_increment_equilibrium=new,
 maximum_equation_change_MPa=float(max(max(abs(newr[:56]-oldr[:56])),source['calibration'][0]*max(abs(newr[56:]-oldr[56:])))),
 six_force_change_N_Nmm=(newf-oldf).tolist(),forward_yield_crossings_on_sampled_segments=crossings,
 no_new_NDS_or_material_acceptance=True,physical_capacity_certified=False,whole_goal_completed=False)
(root/'ColdAcceptedStressContinuityProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('POSTCOMMIT ZERO-STRESS CHANGE',proof['maximum_postcommit_concrete_stress_difference_MPa'],'MPa;EQUATION CHANGE',proof['maximum_equation_change_MPa'])
print('BEFORE',old,'AFTER',new,'YIELD CROSSINGS',len(crossings))
