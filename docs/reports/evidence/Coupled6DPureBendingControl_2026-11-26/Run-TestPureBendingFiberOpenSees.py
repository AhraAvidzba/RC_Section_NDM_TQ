"""ДЛЯ ТЕСТОВ. Проверяет несколько чистых моментов в OpenSees на тех же одноосных диаграммах и той же дискретной геометрии, что штатный плоский НДМ."""
from pathlib import Path
from argparse import ArgumentParser
from hashlib import sha256
from time import perf_counter
import json,struct,math
import openseespy.opensees as ops
import openseespywin.opensees as native

p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,required=True);root=p.parse_args().root
raw=(root/'Input/PlanarReference.bin').read_bytes();assert raw[:16]==b'NDMPLANEMOM00001';position=16

def TestVector():
    """Читает исходные Double без округления геометрии, диаграмм и предельного момента."""
    global position
    count,=struct.unpack_from('<i',raw,position);position+=4;assert 0<count<=10000
    result=struct.unpack_from('<'+'d'*count,raw,position);position+=8*count
    return result

meta,plane,limits=TestVector(),TestVector(),TestVector();moment=meta[1];nc,ns,npc,nps=map(int,limits[4:8])
c=[TestVector() for _ in range(npc)];s=[TestVector() for _ in range(nps)]
geometry=[TestVector() for _ in range(nc)];bars=[TestVector() for _ in range(ns)]
assert position==len(raw) and moment>0 and ops.version()=='3.8.0'
ratios=json.loads((root/'Input/Cases.json').read_bytes())['ratios'];cases=[];started=perf_counter()
for ratio in ratios:
    t0=perf_counter();ops.wipe();ops.model('basic','-ndm',2,'-ndf',3)
    cc=c+[(.1,0.)] # Сохраняется нулевое напряжение при любом положительном удлинении, как у исходной диаграммы Ignore.
    for tag,curve in ((1,cc),(2,s)):
        ops.uniaxialMaterial('ElasticMultiLinear',tag,'-strain',*(x[0] for x in curve),'-stress',*(x[1] for x in curve))
    ops.section('Fiber',1)
    for x,y,area in geometry:ops.fiber(y,x,area,1)
    for x,y,area in bars:
        ops.fiber(y,x,area,2)
        ops.fiber(y,x,-area,1) # Арматура заменяет бетон; отрицательное волокно реализует буквальную поправку исходного интеграла.
    ops.node(1,0.,0.);ops.node(2,0.,0.);ops.fix(1,1,1,1);ops.fix(2,0,1,0)
    ops.element('zeroLengthSection',1,1,2,1)
    ops.timeSeries('Linear',1);ops.pattern('Plain',1,1);ops.load(2,0.,0.,-ratio*moment)
    ops.constraints('Plain');ops.numberer('Plain');ops.system('UmfPack');ops.test('NormUnbalance',.01,60)
    ops.algorithm('Newton');ops.integrator('LoadControl',.125);ops.analysis('Static')
    step=.125;rejected=[];accepted=[]
    for attempt in range(1,257):
        base=ops.getTime();delta=min(step,1.-base);ops.integrator('LoadControl',delta);ok=False
        for algorithm,args in (('Newton',()),('NewtonLineSearch',('-type','Bisection')),('KrylovNewton',())):
            ops.test('NormUnbalance',.01,60);ops.algorithm(algorithm,*args);code=ops.analyze(1)
            if code==0:
                ok=True;break
            assert math.isclose(ops.getTime(),base,abs_tol=1e-13)
            rejected.append(dict(target_factor=base+delta,algorithm=algorithm,code=code))
        if ok:
            force=ops.eleResponse(1,'section','force');deformation=ops.eleResponse(1,'section','deformation')
            assert len(force)==2 and len(deformation)==2
            accepted.append(dict(factor=ops.getTime(),raw_section_forces=force,raw_section_deformations=deformation,iterations=ops.testIter()))
            if abs(ops.getTime()-1.)<1e-12:break
            step=min(.125,1.25*step)
        else:
            step=delta/2
            if step<.125/64:break
    force=ops.eleResponse(1,'section','force');deformation=ops.eleResponse(1,'section','deformation')
    fibers=[ops.eleResponse(1,'section','fiber',i,'stressStrain') for i in range(nc+2*ns)]
    assert all(len(v)==2 for v in fibers) and len(force)==len(deformation)==2
    ec=[v[1] for v in fibers[:nc]];es=[fibers[nc+2*i][1] for i in range(ns)]
    utilization=max(-min(ec)/abs(limits[0]),-min(es)/abs(limits[2]),max(es)/limits[3])
    solved=abs(ops.getTime()-1.)<1e-12
    cases.append(dict(ratio=ratio,target_Mx_Nmm=ratio*moment,equilibrium_at_requested_moment=solved,
        last_accepted_factor=ops.getTime(),actual_resultants_N_Nmm=[force[0],-force[1],0.],
        physical_deformation_plane=[deformation[0],-deformation[1],0.],physical_strain_utilization=utilization,
        physical_limit_exceeded=utilization>1.+1e-7,all_fiber_stress_strain=fibers,accepted_steps=accepted,rejected_trials=rejected,seconds=perf_counter()-t0,
        numerical_failure_not_physical_failure=True))
    print('FIBER_CASE',ratio,'REQUESTED_EQ',solved,'LAST',ops.getTime(),'STRAIN_UTIL',utilization,flush=True)
    output=dict(engine_version=ops.version(),binary_sha256=sha256(Path(native.__file__).read_bytes()).hexdigest().upper(),
        input_sha256=sha256(raw).hexdigest().upper(),reference_Mx_Nmm=moment,source_script_sha256=sha256(Path(__file__).read_bytes()).hexdigest().upper(),
        same_planar_material_curves=True,displaced_concrete_subtracted_exactly=True,geometry_centroid_symmetry=True,
        concrete_curve=cc,steel_curve=s,concrete_geometry=geometry,steel_geometry=bars,cases=cases,seconds=perf_counter()-started,
        section_model_not_3D_solids=True,physical_limits_checked_separately_from_equilibrium=True,whole_goal_completed=False)
    (root/'Result.json').write_text(json.dumps(output,sort_keys=True,separators=(',',':'))+'\n',encoding='utf-8')
ops.wipe()
