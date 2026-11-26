"""ДЛЯ ТЕСТОВ. Проверяет объёмный чистый изгиб отдельно от волоконного эталона и не считает непроверенные уровни доказанным отсутствием равновесия."""
from pathlib import Path
from argparse import ArgumentParser
from hashlib import sha256
import ast,json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');parser=ArgumentParser(description=__doc__);parser.add_argument('--root',type=Path,default=base/'Cold1010PureBendingVolumeExactInput_20261126');root=parser.parse_args().root
prep=json.loads((root/'Preparation.json').read_bytes());d=json.loads((root/'Result.json').read_bytes())
reference=json.loads((root/'PureBendingReferenceProof.json').read_bytes());target=np.array([0.,reference['Mx_ultimate_Nmm'],0.,0.,0.,0.])
assert d['termination'] is not None and reference['planar_reference_physical_limit_verified']
assert sha256((root/'validate_opensees_hardshear.py').read_bytes()).hexdigest().upper()==prep['source_sha256']
np.testing.assert_array_equal(d['section_target'],target);assert prep['section_target']==target.tolist()
assert np.all(np.diff([s['factor'] for s in d['states']])>0)
material=json.loads((base/'Cold1005PureBendingReference_20261126/Input/MaterialContract.json').read_bytes())
E=30000.;nu=.2;G=E/(2*(1+nu));la=E*nu/((1+nu)*(1-2*nu));D=np.diag([2*G]*3+[G]*3);D[:3,:3]+=la;alpha=4/33
tree=ast.parse((here/'Check-TestNMTComparison.py').read_bytes());function=next(n for n in tree.body if isinstance(n,ast.FunctionDef) and n.name=='TestASDFirstStress')
exec(compile(ast.Module(body=[function],type_ignores=[]),'<independent-first-ASD-step>','exec'),globals())
maximum=0.;first_error=0.;tensors=0
for index,state in enumerate(d['states']):
    errors=[float(max(abs((np.array(state[key])-state['factor']*target)/[1,180,180,180,1,1]))) for key in ('base_at_cut','lower_cut','upper_cut')]
    assert max(errors)<=d['cut_check_error_N'] and np.max(abs(np.array(errors)-state['equilibrium_error']))<1e-8
    maximum=max(maximum,max(errors))
    for block in state['cut_fields']:
        stress=np.array(block['stresses']).reshape(8,6)[:,[0,1,2,3,5,4]];strain=np.array(block['strains']).reshape(8,6)[:,[0,1,2,3,5,4]]
        assert np.all(np.isfinite(stress)) and np.all(np.isfinite(strain));tensors+=8
        if index==0:
            error=float(np.max(abs(stress-np.array([TestASDFirstStress(e) for e in strain]))));assert error<1e-8;first_error=max(first_error,error)
    for kind,bar in state['steel'].items():
        limit=(.025 if bar['strain']>=0 else .018) if kind.startswith('long') else .02
        assert abs(bar['util']-abs(bar['strain'])/limit)<1e-12
last=d['states'][-1];rows=[]
for ratio in prep['requested_factors']:
    matches=[s for s in d['states'] if abs(s['factor']-ratio)<1e-12]
    rows.append(dict(ratio=ratio,target_Mx_kNm=ratio*reference['Mx_ultimate_kNm'],
        status='equilibrium_found' if matches else 'not_reached_by_this_continuation',equilibrium_absence_not_proved=not bool(matches)))
proof=dict(reference_Mx_kNm=reference['Mx_ultimate_kNm'],last_equilibrium_ratio=last['factor'],
    last_equilibrium_Mx_kNm=last['factor']*reference['Mx_ultimate_kNm'],solver_seconds=d['elapsed_s'],termination=d['termination'],
    maximum_independent_resultant_error_N=maximum,checked_tensors=tensors,first_load_material_error_MPa=first_error,
    last_cut_damage=last['adjacent_damage'],last_maximum_damage=last['prism_damage_max'],
    material_law_not_identical_to_planar_reference=True,transverse_reinforcement_and_strains_present=True,
    auxiliary_elastic_regions_not_homogeneous_beam=True,numerical_stop_not_physical_failure=True,
    requested_cases=rows,physical_capacity_certified=False,whole_goal_completed=False)
(root/'PureBendingVolumeProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PURE BENDING VOLUME VERIFIED',json.dumps({k:v for k,v in proof.items() if k!='requested_cases'},sort_keys=True))
