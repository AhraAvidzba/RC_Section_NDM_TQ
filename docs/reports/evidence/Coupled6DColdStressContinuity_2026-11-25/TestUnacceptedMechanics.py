"""ДЛЯ ТЕСТОВ. Общая независимая сборка непринятых HardShear-корней из физических тензоров, операторов и точных Double."""
from pathlib import Path
from runpy import run_path
import numpy as np

def TestUnacceptedRoots(root,source,cases,phase='EnvelopeEnergy.Accepted',stress_file='ElasticPrincipal.Roots.csv'):
    """Проверяет все силовые строки, шесть усилий, связи и оба класса стали; временного принятия материала не требует."""
    here=Path(__file__).resolve().parent
    rows=run_path(str(here/'Validate-TestHardPrepared.py'))['TestRows'](root/'Native'/stress_file)
    g,L,w,v,coords,axis,b,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native'/(phase+'.Operators.csv'),60,43)
    driver=(root/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig')
    assert 'TestInputCell(steel, "Steel.TwoLine.Es2").Value2 = 0.018' in driver
    assert 'TestInputCell(steel, "Steel.TwoLine.Es2", 1, 3).Value2 = 0.025' in driver
    assert 'grades.Cells(2, 6).Value2 = 0.02: grades.Cells(2, 7).Value2 = 0.02' in driver
    assert set(axis[25:,2])=={0.,1.}
    stress={};bar={};energy={}
    for row in rows:
        kind,case,point,side=row[:4];key=(int(case),int(side),int(point))
        if kind=='K':stress[key]=np.array(row[4:10],float);bar[key]=float(row[10])
        elif kind=='E':energy[key]=float(row[4])
        else:assert kind=='P',kind
    weight=coords[:,2];center=np.average(coords[:25,:2],axis=0,weights=weight[:25]);x,y=(coords[:,:2]-center).T
    scales=np.array([1/L[0,0],1/L[1,1],1/L[1,1],1/L[1,1],1/L[0,0],1/L[0,0]])
    op=np.array([b+s*.001*gradient for s in (-1,0,1)])
    M=np.einsum('spcd,spce,p,c->de',op,op,weight,np.array([1.,1.,1.,.5,.5,.5]))/(3*sum(weight))
    M[56:,:]=0;M[:,56:]=0
    results={}
    for case,item in cases.items():
        field=item['field'];factor=item['lambda_value'];accuracy=item['accuracy'];net=np.zeros((3,43,6));utilization=[]
        for side in (-1,0,1):
            maximum=0.
            for point in range(1,44):
                s=stress[case,side,point];net[side+1,point-1]=s if point<=25 else bar[case,side,point]*axis[point-1]-s
                if point>25:
                    axial=float((op[side+1,point-1]@field)@axis[point-1]);limit=(.025 if axial>=0 else .018) if axis[point-1,2]==1 else .02
                    maximum=max(maximum,abs(axial)/limit)
            utilization.append(maximum)
        assert max(utilization)<=1
        target=source['offset']+factor*source['direction'];external=L@target
        weak=g@field+np.einsum('pcd,pc->d',w,net[1])+np.einsum('pcd,pc->d',v,(net[2]-net[0])/.002)-external
        weak_norm=max(float(max(abs(weak[:56]))),float(source['calibration'][0]*max(abs(weak[56:]))))
        s=net[1];actual=np.array([weight@s[:,2],(weight*y)@s[:,2],(weight*x)@s[:,2],weight@(x*s[:,5]-y*s[:,4]),weight@s[:,4],weight@s[:,5]])
        force_norm=float(max(abs((actual-target)/scales)))
        bound=float(accuracy*(source['options'][3]+source['options'][4]*max(abs(external[:56]))))
        assert weak_norm<=bound+1e-12 and force_norm<=bound+1e-12,(case,weak_norm,force_norm,bound)
        if 'actual' in item:assert max(abs((item['actual']-target)/scales))<=bound
        h=None
        if energy:
            h=sum(weight[point-1]*energy[case,side,point] for point in range(1,26) for side in (-1,0,1))/(3*sum(weight[:25]))
        results[case]=dict(weak_norm_MPa=weak_norm,six_force_norm_MPa=force_norm,bound_MPa=bound,
            steel_utilization_minus_center_plus=utilization,independent_point_energy_average_MPa=h)
    assert len(stress)==129*len(cases)
    return results,M
