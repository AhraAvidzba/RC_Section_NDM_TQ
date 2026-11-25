"""ДЛЯ ТЕСТОВ. Независимо сравнивает локальные касательные с прямыми разностями и собирает их действие в56+4 уравнениях."""
from pathlib import Path
from runpy import run_path
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');root=base/'Cold872PointDerivativeAudit20_20261126';parent=base/'Cold867FixedMetricKernel_20261126'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=2,failed=0,macros=2,timeout600=False,own_process_closed=True)
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes());assert audit['components']==256 and audit['imported_this_pass']==['modTestCoupledElasticState']
assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==255
for f in (parent/'src').rglob('*'):
    if f.is_file():assert f.read_bytes()==(root/f.relative_to(parent)).read_bytes(),f
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin');assert len(source['fields'])==20
assert (root/'Input/Reverse1.bin').read_bytes()==(base/'Cold869FixedMetricNextScan_20261126/Input/Reverse1.bin').read_bytes()
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)]);E=source['calibration'][0]
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((root/'Native/PointDerivative.bin').read_bytes());assert r.raw[:16]==b'NDMPOINTDERIV001'
header=r.VectorTest();probe=r.VectorTest();movement=r.VectorTest();assert tuple(header[:2])==(2,43)
np.testing.assert_array_equal(probe,source['fields'][-1]+(source['fields'][-1]-source['fields'][-2])/1024.)
guide=R((parent/'Native/LoadingWalk.Guide003.bin').read_bytes());guide.VectorTest()
for size in (3,3):
    if guide.VectorTest()[0]:
        for _ in range(size):guide.VectorTest()
assert guide.VectorTest()[0]==1
tm,_,_,expected=[guide.VectorTest() for _ in range(4)];np.testing.assert_array_equal(movement,expected)
records={};fields={};local=[];steps=(0.,.0001,.001,.01,.1)
kinematic_roundoff=[]
def TestSerializedKinematics(actual,field,side,point):
    """Учитывает округление коэффициентов CSV до15 значащих цифр и суммирование60 членов; не меняет расчётные допуски сил."""
    expected=op[side+1,point-1]@field
    absolute_sum=(abs(B[point-1])+.001*abs(gradient[point-1]))@abs(field)
    bound=(8e-15+64*np.finfo(float).eps)*absolute_sum+np.finfo(float).tiny
    defect=abs(actual-expected)
    assert np.all(defect<=bound),(side,point,defect,bound)
    kinematic_roundoff.append(float(max(defect)))
for location in (1,2):
    field=r.VectorTest();fields[location]=field;np.testing.assert_array_equal(field,source['fields'][-1] if location==1 else probe)
    for side in (-1,0,1):
        for point in range(1,44):
            info=r.VectorTest();assert tuple(info)==(location,side,point)
            strain=r.VectorTest();direction=r.VectorTest()
            TestSerializedKinematics(strain,field,side,point)
            TestSerializedKinematics(direction,movement,side,point)
            modes={}
            for mode in (0,1):
                mi=r.VectorTest();assert mi[0]==mode
                modes[mode]=None if mi[1] else dict(stress=r.VectorTest(),D=np.array([r.VectorTest() for _ in range(6)]).T)
            direct={}
            for h in steps:
                for sign in (-1,1):
                    di=r.VectorTest();assert di[0]==h and di[1]==sign and di[3] in (0,1)
                    direct[h,sign]=dict(info=di,stress=r.VectorTest() if di[2]==0 else None)
            record=dict(strain=strain,direction=direction,modes=modes,direct=direct);records[location,side,point]=record
            base_stress=direct[0.,1]['stress'];assert base_stress is not None
            np.testing.assert_array_equal(base_stress,direct[0.,-1]['stress'])
            for mode,item in modes.items():
                if item is None:continue
                np.testing.assert_array_equal(item['stress'],base_stress)
                predicted=item['D']@direction
                for h in steps[1:]:
                    plus,minus=direct[h,1],direct[h,-1]
                    if plus['stress'] is None or minus['stress'] is None:continue
                    actual=(plus['stress']-minus['stress'])/(2*h)
                    local.append(dict(location=location,side=side,point=point,mode=mode,h=h,
                        error_MPa=float(max(abs(actual-predicted))),relative_E_strain_error=float(max(abs(actual-predicted))/max(E*max(abs(direction)),1e-30)),
                        plus_branch=plus['info'][4:].tolist(),minus_branch=minus['info'][4:].tolist(),
                        base_branch=direct[0.,1]['info'][4:].tolist()))
assert r.position==len(r.raw)

def TestNorm(vector):
    """Сохраняет масштаб МПа физических уравнений и исходное масштабирование четырёх связей."""
    return float(max(max(abs(vector[:56])),E*max(abs(vector[56:]))))

def TestDerivativeFromStress(location,h):
    """Собирает прямую центральную производную бетонных напряжений с аналитически упругой сталью замещения."""
    ds=np.zeros((3,43,6));bar_direction=np.einsum('spc,pc->sp',np.einsum('spcd,d->spc',op,movement)[:,25:],axis[25:])
    for sign in (-1,1):
        actual_strain=np.einsum('spcd,d->spc',op,fields[location]+sign*h*movement)
        bar_stress=200000*np.einsum('spc,pc->sp',actual_strain[:,25:],axis[25:])
        assert np.max(abs(bar_stress))<240
    for side in (-1,0,1):
        for point in range(1,44):
            d=records[location,side,point]['direct'];plus=d[h,1]['stress'];minus=d[h,-1]['stress']
            assert plus is not None and minus is not None
            ds[side+1,point-1]=(plus-minus)/(2*h)
    ds[:,25:]=200000*bar_direction[:,:,None]*axis[None,25:]-ds[:,25:]
    return g@movement+np.einsum('pcd,pc->d',w,ds[1])+np.einsum('pcd,pc->d',v,(ds[2]-ds[0])/.002)-L@source['direction']*tm[1]

global_checks=[]
for location in (1,2):
    direct_vectors={h:TestDerivativeFromStress(location,h) for h in steps[1:]}
    for mode in (0,1):
        errors=[key for key,record in records.items() if key[0]==location and record['modes'][mode] is None]
        item=dict(location=location,mode=mode,unavailable_points=[list(k) for k in errors])
        if not errors:
            ds=np.array([[records[location,side,point]['modes'][mode]['D']@records[location,side,point]['direction'] for point in range(1,44)] for side in (-1,0,1)])
            steel=np.einsum('spc,pc->sp',np.einsum('spcd,d->spc',op,movement)[:,25:],axis[25:])
            ds[:,25:]=200000*steel[:,:,None]*axis[None,25:]-ds[:,25:]
            assembled=g@movement+np.einsum('pcd,pc->d',w,ds[1])+np.einsum('pcd,pc->d',v,(ds[2]-ds[0])/.002)-L@source['direction']*tm[1]
            item.update(assembled_residual_derivative_MPa=TestNorm(assembled),direct=[])
            for h,direct in direct_vectors.items():
                defect=direct-assembled
                item['direct'].append(dict(h=h,direct_derivative_norm_MPa=TestNorm(direct),defect_MPa=TestNorm(defect),dominant_physical_row=int(np.argmax(abs(defect[:56])))+1))
        global_checks.append(item)
proof=dict(native=native,actual_components=256,all_working_material_methods_exact867=True,source20_exact=True,
    maximum_serialized_operator_kinematic_roundoff=float(max(kinematic_roundoff)),
    no_new_NDS_or_history_acceptance=True,global_checks=global_checks,local_comparisons=local,
    local_largest_errors=sorted(local,key=lambda x:x['error_MPa'],reverse=True)[:30],physical_capacity_certified=False)
(root/'PointDerivativeAuditProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
for item in global_checks:print('LOCATION/MODE',item['location'],item['mode'],'UNAVAILABLE',item['unavailable_points'],'ASSEMBLED',item.get('assembled_residual_derivative_MPa'),'DIRECT',item.get('direct'))
print('LARGEST LOCAL',proof['local_largest_errors'][:4])
