"""ДЛЯ ТЕСТОВ. Независимо проверяет все волокна, знаки, физические ограничения и найденные/не найденные чистые моменты OpenSees."""
from pathlib import Path
from hashlib import sha256
import json,numpy as np

root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1007PureBendingFiberOpenSees_20261126')
prep=json.loads((root/'Preparation.json').read_bytes());d=json.loads((root/'Result.json').read_bytes())
for name,digest in prep['inputs'].items():assert sha256((root/name).read_bytes()).hexdigest().upper()==digest
assert d['source_script_sha256']==prep['inputs']['Run-TestPureBendingFiberOpenSees.py'] and d['engine_version']=='3.8.0'
assert d['binary_sha256']=='25BA42E1672526137D765D9C3631D1403E091F54F5AFCDB9242F9243C63C3A96'
reference=json.loads((root/'Input/PureBendingReferenceProof.json').read_bytes());assert reference['planar_reference_physical_limit_verified']
assert d['input_sha256']==prep['inputs']['Input/PlanarReference.bin'] and d['reference_Mx_Nmm']==reference['Mx_ultimate_Nmm']
assert [c['ratio'] for c in d['cases']]==prep['selected_ratios']
concrete=np.array(d['concrete_curve']);steel=np.array(d['steel_curve']);geometry=np.array(d['concrete_geometry']);bars=np.array(d['steel_geometry'])
nc,ns=len(geometry),len(bars);max_error=0.;cases=[]
for case in d['cases']:
    assert case['target_Mx_Nmm']==case['ratio']*reference['Mx_ultimate_Nmm']
    eps,kx,ky=case['physical_deformation_plane'];assert ky==0
    ec=eps+kx*geometry[:,1];es=eps+kx*bars[:,1]
    sc=np.interp(ec,concrete[:,0],concrete[:,1]);ss=np.interp(es,steel[:,0],steel[:,1]);cs=np.interp(es,concrete[:,0],concrete[:,1])
    actual=np.array([geometry[:,2]@sc+bars[:,2]@(ss-cs),(geometry[:,2]*geometry[:,1])@sc+(bars[:,2]*bars[:,1])@(ss-cs),0.])
    observed=np.array(case['actual_resultants_N_Nmm']);error=float(max(abs(actual-observed)));assert error<1e-5;max_error=max(max_error,error)
    assert abs(actual[0])<.01 and abs(actual[1]-case['last_accepted_factor']*case['target_Mx_Nmm'])<.01
    rows=np.array(case['all_fiber_stress_strain']);expected=[]
    for stress,strain in zip(sc,ec):expected.append([stress,strain])
    for i in range(ns):expected.extend([[ss[i],es[i]],[cs[i],es[i]]])
    assert rows.shape==(nc+2*ns,2);np.testing.assert_allclose(rows,np.array(expected),atol=1e-8,rtol=0)
    utilization=max(-min(ec)/abs(concrete[0,0]),-min(es)/abs(steel[0,0]),max(es)/steel[-1,0])
    assert abs(utilization-case['physical_strain_utilization'])<1e-9
    assert case['equilibrium_at_requested_moment']==(abs(case['last_accepted_factor']-1.)<1e-12)
    assert case['physical_limit_exceeded']==(utilization>1.+1e-7)
    cases.append(dict(ratio=case['ratio'],target_Mx_kNm=case['target_Mx_Nmm']/1e6,equilibrium_at_requested_moment=case['equilibrium_at_requested_moment'],
        last_equilibrium_Mx_kNm=actual[1]/1e6,last_equilibrium_ratio=actual[1]/reference['Mx_ultimate_Nmm'],
        strain_utilization=utilization,physical_limit_exceeded=case['physical_limit_exceeded'],seconds=case['seconds']))
at_limit=next(c for c in cases if c['ratio']==1.);assert at_limit['equilibrium_at_requested_moment'] and abs(at_limit['strain_utilization']-1.)<1e-6
proof=dict(reference_Mx_kNm=reference['Mx_ultimate_kNm'],maximum_requested_ratio_with_equilibrium=max(c['ratio'] for c in cases if c['equilibrium_at_requested_moment']),
    maximum_retained_equilibrium_ratio=max(c['last_equilibrium_ratio'] for c in cases),maximum_independent_resultant_error_N_Nmm=max_error,
    every_fiber_stress_strain_verified=True,material_curves_exact_planar_reference=True,elapsed_seconds=d['seconds'],cases=cases,
    physical_limits_separate_from_equilibrium=True,OpenSees_section_not_3D_solids=True,whole_goal_completed=False)
(root/'PureBendingFiberProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PURE BENDING FIBER VERIFIED',json.dumps({k:v for k,v in proof.items() if k!='cases'},sort_keys=True))
