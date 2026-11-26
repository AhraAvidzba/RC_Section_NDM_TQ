"""ДЛЯ ТЕСТОВ. Проверяет инженерные объёмные контроли, локальные тензоры, равновесие и чувствительность к удалённости/закреплениям торцов."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
import json,numpy as np

base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
names={993:'Cold993OpenSeesMatchedForce_20261126',994:'Cold994OpenSeesElasticBoundary_20261126',995:'Cold995OpenSeesElasticBoundary_20261126',
    996:'Cold996OpenSeesElasticBoundary_20261126',997:'Cold997OpenSeesElasticClamped_20261126',998:'Cold998OpenSeesBufferedCore_20261126',999:'Cold999OpenSeesBufferedCore_20261126'}
target=np.array([28000.,510000.,180000.,1500000.,17000.,-7100.]);scale=np.array([1,180,180,180,1,1])
cases={};raw={}
for n,name in names.items():
    root=base/name
    if not (root/'Result.json').is_file():continue
    prep=json.loads((root/'Preparation.json').read_bytes());d=json.loads((root/'Result.json').read_bytes())
    assert d['termination'] is not None,'Running process is not a completed result'
    assert sha256((root/'validate_opensees_hardshear.py').read_bytes()).hexdigest().upper()==prep['source_sha256']
    np.testing.assert_array_equal(d['section_target'],target)
    assert len(d['states'])>0 and np.all(np.diff([s['factor'] for s in d['states']])>0)
    maximum_error=0.;elastic_error=0.;tensor_count=0
    for state in d['states']:
        force=state['factor']*target
        errors=[float(max(abs((np.asarray(state[key])-force)/scale))) for key in ('base_at_cut','lower_cut','upper_cut')]
        assert max(errors)<=d['cut_check_error_N'] and max(abs(np.array(errors)-state['equilibrium_error']))<1e-8
        maximum_error=max(maximum_error,max(errors))
        for kind,bar in state['steel'].items():
            limit=(.025 if bar['strain']>=0 else .018) if kind.startswith('long') else .02
            assert abs(bar['util']-abs(bar['strain'])/limit)<1e-12
        if d.get('nonlinear_zone_mm'):
            lo,hi=d['nonlinear_zone_mm']
            for item in state['prism_damage_max']:
                if item['value']>0:assert lo<=item['z_interval_mm'][0] and item['z_interval_mm'][1]<=hi
        for block in state.get('cut_fields',[]):
            stresses=np.asarray(block['stresses']).reshape(8,6);strains=np.asarray(block['strains']).reshape(8,6)
            assert np.all(np.isfinite(stresses)) and np.all(np.isfinite(strains));tensor_count+=8
            if d['variant']=='elastic':
                E,nu=30000.,.2;G=E/(2*(1+nu));lame=E*nu/((1+nu)*(1-2*nu))
                predicted=G*strains;predicted[:,:3]=2*G*strains[:,:3]+lame*np.sum(strains[:,:3],axis=1)[:,None]
                relative=float(max(abs((predicted-stresses).ravel()))/max(1.,max(abs(stresses.ravel()))))
                assert relative<1e-10;elastic_error=max(elastic_error,relative)
    if d['variant']=='matched':assert max(max(a) for a in d['installed_adjustment'])<1e-7
    run=json.loads((root/'Run.ndjson').read_text('utf-8-sig').splitlines()[-1]);start=datetime.fromisoformat(run['utc'].replace('Z','+00:00'))
    finish=datetime.fromtimestamp((root/'Result.json').stat().st_mtime,start.tzinfo)
    wanted=prep.get('requested_factors',[1.] if d['variant']=='elastic' else [])
    item=dict(length_mm=d['length'],variant=d['variant'],end_condition=d['end_condition'],mesh=d['mesh'],
        nonlinear_zone_mm=d.get('nonlinear_zone_mm'),accepted_states=len(d['states']),rejected_trials=len(d['rejected']),
        maximum_accepted_lambda=max(s['factor'] for s in d['states']),termination=d['termination'],
        analysis_seconds=d['elapsed_s'],approximate_full_run_seconds=(finish-start).total_seconds(),
        maximum_independent_cut_force_error_N=maximum_error,checked_stress_strain_tensors=tensor_count,
        maximum_elastic_constitutive_relative_error=elastic_error,last_cut_damage=d['states'][-1]['adjacent_damage'],
        last_prism_damage=d['states'][-1]['prism_damage_max'],
        all_requested_force_marks_reached=bool(wanted) and all(any(abs(s['factor']-mark)<1e-10 for s in d['states']) for mark in wanted),
        numerical_stop_not_physical_capacity=True,spatial_law_not_identical_NDM_CDP=d['variant']!='elastic',physical_capacity_certified=False)
    (root/'EngineeringOpenSeesProof.json').write_text(json.dumps(item,sort_keys=True,indent=2)+'\n',encoding='utf-8')
    cases[str(n)]=item;raw[n]=d

def TestCompare(a,b):
    """Сопоставляет одинаковые точки Гаусса центральных слоёв при точно совпавшей нагрузке; не подменяет нелинейный контроль упругим."""
    rows=[]
    for first in raw[a]['states']:
        candidates=[s for s in raw[b]['states'] if abs(s['factor']-first['factor'])<1e-12]
        if not candidates or not first.get('cut_fields'):continue
        second=candidates[0];x=first['cut_fields'];y=second['cut_fields']
        assert [r['center_relative_to_cut_mm'] for r in x]==[r['center_relative_to_cut_mm'] for r in y]
        result=dict(lambda_value=first['factor'])
        for key in ('stresses','strains'):
            p=np.array([r[key] for r in x]);q=np.array([r[key] for r in y])
            result[key+'_difference_percent']=float(np.max(abs(p-q))/max(np.max(abs(q)),np.finfo(float).tiny)*100)
        rows.append(result)
    return dict(first=a,second=b,compared_common_load_states=len(rows),states=rows,
        both_runs_reach_requested_maximum=cases[str(a)]['all_requested_force_marks_reached'] and cases[str(b)]['all_requested_force_marks_reached'])

comparisons=[TestCompare(a,b) for a,b in ((994,996),(995,996),(996,997),(998,999)) if a in raw and b in raw]
proof=dict(cases=cases,comparisons=comparisons,no_PhysicalFail_from_numerical_stop=True,
    homogeneous_prism_end_failure_not_section_capacity=True,buffer_regions_are_explicit_engineering_assumption=True,
    physical_capacity_certified=False,whole_goal_completed=False)
(base/'EngineeringOpenSeesBoundaryProof_20261126.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('ENGINEERING OPENSEES VERIFIED',json.dumps(proof,sort_keys=True))
