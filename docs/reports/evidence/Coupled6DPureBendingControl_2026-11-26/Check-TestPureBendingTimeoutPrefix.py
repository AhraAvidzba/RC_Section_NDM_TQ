"""ДЛЯ ТЕСТОВ. Проверяет завершённый префикс CDP-набора с тайм-аутом, не присваивая невыполненным моментам NumFail или физический отказ."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np

here=Path(__file__).resolve().parent;root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1006PureBendingCoupled_20261126')
closure=json.loads((root/'Native/RealColdReplay.txt.ProcessClosure.json').read_text('utf-8-sig'))
owner=json.loads((root/'Native/RealColdReplay.txt.Ownership.json').read_text('utf-8-sig'))
assert closure['Stage']=='TIMEOUT' and not closure['Passed'] and closure['After']['State']=='Absent'
assert closure['Identity']['Valid'] and closure['Identity']['PID']==owner['PID'] and int(closure['Identity']['Ticks'])==int(owner['StartTicksUTC'])
compiled=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']==owner['SourceSHA256']
report=(root/'Native/RealColdReplay.txt').read_text('utf-8-sig')
assert 'WATCHDOG_FAILURE: timeout=600 sec' in report and 'ownExcelState=Absent; exactCleanupRequested=True' in report
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((root/'Native/PureBendingCoupled.bin').read_bytes())
assert r.raw[:16]==b'NDMPURECOUP00001';header=r.VectorTest();mu=header[0];assert header[1:6].tolist()==[2.,11.,60.,43.,25.]
meta=r.VectorTest();assert meta[0]==1 and meta[1]==.25 and meta[3]==1 and meta[4]==1
field=r.VectorTest();reported=r.VectorTest();strain=np.zeros((3,43,6));net=strain.copy();sigma=strain.copy()
for k in range(129):
    point,side,code=r.VectorTest();assert point==k%43+1 and side==k//43-1 and code==0
    strain[k//43,k%43]=r.VectorTest();net[k//43,k%43]=r.VectorTest();sigma[k//43,k%43]=r.VectorTest()
history=[]
for _ in range(int(meta[-1])):r.VectorTest();history.append(r.VectorTest())
assert r.position==len(r.raw),'Partial trailing case must be reported explicitly'
np.testing.assert_array_equal(history[-1],field)
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/PureBending.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)])
np.testing.assert_allclose(strain,np.einsum('spcd,d->spc',op,field),atol=1e-14,rtol=0)
eps=np.einsum('spc,pc->sp',strain[:,25:],axis[25:]);long=axis[25:,2]==1;fy=np.where(long[None,:],np.where(eps>=0,400.,350.),240.)
assert np.all(abs(200000.*eps)<fy)
assert np.max(abs(net[:,25:]+sigma[:,25:]-(200000.*eps)[:,:,None]*axis[None,25:]))<1e-8
external=L@np.array([0.,mu*.25,0.,0.,0.,0.]);residual=g@field-external+np.einsum('pcd,pc->d',w,net[1])+np.einsum('pcd,pc->d',v,(net[2]-net[0])/.002)
area=coords[:,2];center=np.average(coords[:25,:2],axis=0,weights=area[:25]);x,y=(coords[:,:2]-center).T
actual=np.array([area@net[1,:,2],(area*y)@net[1,:,2],(area*x)@net[1,:,2],area@(x*net[1,:,5]-y*net[1,:,4]),area@net[1,:,4],area@net[1,:,5]])
scales=np.array([1/L[0,0],1/L[1,1],1/L[1,1],1/L[1,1],1/L[0,0],1/L[0,0]])
bound=float(header[7]+header[8]*max(abs(external[:56])));norm=max(float(max(abs(residual[:56]))),float(30000.*max(abs(residual[56:]))),float(max(abs((actual-[0.,mu*.25,0.,0.,0.,0.])/scales))))
assert norm<=bound+1e-10 and max(abs((actual-reported)/scales))<1e-10
proof=dict(planned_cases=22,completed_cases=1,completed_CDP_ratio=.25,completed_Mx_kNm=actual[1]/1e6,completed_equilibrium_verified=True,
    completed_residual_MPa=norm,original_bound_MPa=bound,watchdog_timeout_seconds=600,active_case_at_timeout=dict(model='CDP',ratio=.5),
    active_case_final_equilibrium_unknown=True,remaining_cases_not_executed=True,whole_run_passed=False,own_exact_excel_closed=True,
    source_unchanged=True,no_NumFail_or_PhysicalFail_assigned_to_timeout=True,physical_capacity_certified=False,whole_goal_completed=False)
(root/'PureBendingTimeoutPrefixProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('TIMEOUT PREFIX VERIFIED',json.dumps(proof,sort_keys=True))
