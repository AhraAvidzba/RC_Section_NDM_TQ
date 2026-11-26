"""ДЛЯ ТЕСТОВ. Пересобирает однократное спектральное обновление по первичному исходнику и разделяет ошибку VBA, библиотеки и численного нижнего уровня таблицы."""
from pathlib import Path
import json,numpy as np

root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1001ASDNativeUpdate_20261126')
data=json.loads((root/'Native/OpenSeesNativeMaterialReplay.json').read_bytes())
contract=json.loads((root/'Input/MaterialContract.json').read_bytes());E=30000.;nu=.2;G=E/(2*(1+nu));la=E*nu/((1+nu)*(1-2*nu))
D=np.diag([2*G]*3+[G]*3);D[:3,:3]+=la;alpha=4/33;ratio=10.;beta=(1-alpha)*ratio-(1+alpha)
curves={key:np.array(rows) for key,rows in contract['uniaxial_tables'].items()}
proof={}
for floor_mode in ('prepared_native','OpenSees_source'):
    floors={key:(min(rows[1:,1]) if floor_mode=='prepared_native' else min(abs(v) for v in np.diff(rows[:,1]) if v!=0))*1e-6 for key,rows in curves.items()}
    def TestBranch(kind,x):
        """Сохраняет допуск выбора отрезка источника; ординаты интерполирует в номинальном и эффективном пространстве."""
        rows=curves[kind];xs=rows[:,0];nom=rows[:,1];eff=nom/(1-rows[:,2]);tol=min(np.diff(xs))*1e-6
        i=min(int(np.searchsorted(xs+tol,x,side='left')),len(xs)-1);i=max(1,i)
        if x<=xs[-1]+tol:
            f=(x-xs[i-1])/(xs[i]-xs[i-1]);n=nom[i-1]+f*(nom[i]-nom[i-1]);q=eff[i-1]+f*(eff[i]-eff[i-1])
        else:
            f=(x-xs[-1])/(xs[-1]-xs[-2]);n=nom[-1]+f*max(0,nom[-1]-nom[-2]);q=eff[-1]+f*max(0,eff[-1]-eff[-2])
        return max(floors[kind],n),max(floors[kind],q)
    count=len(data['strain_histories'][0]);old_eps=np.zeros((count,6));old_eff=old_eps.copy();old_coord=np.zeros((count,2));observed=[]
    for frame in data['strain_histories']:
        frame=np.array(frame);current=[]
        for p,eps in enumerate(frame):
            trial=old_eff[p]+D@(eps-old_eps[p]);xx,yy,zz,xy,xz,yz=trial
            tensor=np.array([[xx,xy,xz],[xy,yy,yz],[xz,yz,zz]])
            eig,axes=np.linalg.eigh(tensor);trace=sum(eig);q=np.sqrt(1.5*sum((eig-trace/3)**2));pos=eig[-1]
            negative=np.minimum(eig,0);ntrace=sum(negative);nq=np.sqrt(1.5*sum((negative-ntrace/3)**2))
            excursion=[(alpha*trace+q+beta*pos)/((1-alpha)*E*ratio) if pos>=floors['T'] else 0.,
                       (alpha*ntrace+nq+3*negative[-1])/((1-alpha)*E) if eig[0]<0 else 0.]
            scales=[];nominal=[]
            for k,kind in enumerate(('T','C')):
                prior=old_coord[p,k];n,q=TestBranch(kind,prior);pl=prior-q/E
                now=max(prior,excursion[k]+pl);n,q=TestBranch(kind,now);den=E*(now-pl)
                scales.append(q/den);nominal.append(n/den);old_coord[p,k]=now
            def TestTensor(scales):
                """Спектрально собирает знаковые части без полиномиальной формулы проекторов VBA."""
                t=(axes*(eig*np.where(eig>0,scales[0],scales[1])))@axes.T
                return np.array([t[0,0],t[1,1],t[2,2],t[0,1],t[0,2],t[1,2]])
            old_eff[p]=TestTensor(scales);current.append(TestTensor(nominal))
        old_eps=frame.copy();observed.append(current)
    observed=np.array(observed);native=np.array(data['expected_VBA_concrete_stresses']);external=np.array(data['observed_OpenSees_concrete_stresses'])
    item=dict(native_error_MPa=float(np.max(abs(observed-native))),OpenSees_error_MPa=float(np.max(abs(observed-external))),floors=floors)
    for key,array in (('native_largest',abs(observed-native)),('OpenSees_largest',abs(observed-external))):
        index=np.unravel_index(np.argmax(array),array.shape);f,p,c=map(int,index)
        item[key]=dict(frame=f+1,point=p%43+1,side=p//43-1,component=c+1,
                       source_stress=observed[f,p].tolist(),native_stress=native[f,p].tolist(),OpenSees_stress=external[f,p].tolist())
    proof[floor_mode]=item
(root/'Native/ASDNativeSourceComparison.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(json.dumps(proof,sort_keys=True))
