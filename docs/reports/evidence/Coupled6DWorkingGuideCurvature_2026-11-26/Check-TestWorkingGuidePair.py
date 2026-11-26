"""ДЛЯ ТЕСТОВ. Проверяет пассивность записи949 и независимо восстанавливает равновесие и кривизну фактической полной/половинной пары."""
from pathlib import Path
from runpy import run_path
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold949WorkingGuidePair_20261126';parent=base/'Cold947WorkingSecantReturn_20261126'
script=here/'Check-TestWorkingSecantReturn.py';code=script.read_text('utf-8-sig')
assert code.count('Cold947WorkingSecantReturn_20261126')==2
code=code.replace('Cold947WorkingSecantReturn_20261126','Cold949WorkingGuidePair_20261126')
scope={'__file__':str(script),'__name__':'passive_guide_pair_test'};exec(compile(code,str(script),'exec'),scope)
assert (root/'Native/WorkingGuideReturn.bin').read_bytes()==(parent/'Native/WorkingGuideReturn.bin').read_bytes()
inner=scope['scope'];R=inner['R'];op=inner['op'];coords=inner['coords'];source=inner['source'];meta=inner['meta'];L=inner['L']
reference=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
def TestNorm(field):
    """Вычисляет среднюю тензорную норму по трём станциям независимо от компактного девятистрочного оператора VBA."""
    strain=np.einsum('spcd,d->spc',op,field)
    return float(np.sqrt(np.einsum('spc,spc,p,c->',strain,strain,coords[:,2],np.array([1,1,1,.5,.5,.5]))/(3*sum(coords[:,2]))))
r=R((root/'Native/WorkingGuidePairs.bin').read_bytes());assert r.raw[:16]==b'NDMGUIDEPAIR0001'
count=int(r.VectorTest()[0]);assert count==1;pairs=[]
for index in range(count):
    info,accepted,full,half,projection=[r.VectorTest() for _ in range(5)]
    assert len(info)==13 and all(len(a)==60 for a in (accepted,full,half,projection))
    np.testing.assert_array_equal(accepted,source['fields'][-1])
    full_lam,half_lam,reference_lam,target,half_target,tolerance,native_field_error,native_lambda_error,native_lambda_bound,native_length,alpha,native_joint,native_step=info
    assert reference_lam==reference['lambdas'][-1] and target==meta[4]
    entries=[]
    for name,field,lam,goal in [('full',full,full_lam,target),('half',half,half_lam,half_target)]:
        stress=np.zeros((3,43,6));seen=set()
        for _ in range(129):
            point,side,error=r.VectorTest();sigma=r.VectorTest();key=int(point),int(side)
            assert error==0 and key not in seen and len(sigma)==6;seen.add(key);stress[int(side)+1,int(point)-1]=sigma
        assembled=inner['TestAssemble'](field,lam,stress)
        control=float(meta[8]*(projection@field-goal))
        gauge=float(meta[8]*max(abs(assembled['residual'][56:])))
        norm=float(max(assembled['physical_norm'],assembled['six_force_norm'],gauge,abs(control)))
        target_load=source['offset']+lam*source['direction'];bound=float(source['options'][3]+source['options'][4]*max(abs((L@target_load)[:56])))
        assert norm<=bound+2e-11,(name,norm,bound)
        strain=np.einsum('spcd,d->spc',op,field);bars=np.einsum('spc,pc->sp',strain[:,25:],inner['axis'][25:]);long=inner['axis'][25:,2]==1
        steel_util=float(np.max(abs(bars)/np.where(long[None,:],np.where(bars>=0,.025,.018),.02)))
        assert steel_util<=1
        branches=[]
        for point in range(1,44):
            for side in (-1,0,1):
                value=inner['TestYield'](point,side,field);threshold=1e-11*inner['states'][point,side][8]
                branches.append(dict(point=point,side=side,real_concrete=point<=25,trial_yield_MPa=value,
                    branch=('plastic_loading' if value>threshold else 'elastic_unloading' if value < -threshold else 'boundary')))
        entries.append(dict(name=name,lambda_value=float(lam),target=float(goal),independent_total_residual_MPa=norm,
            final_bound_MPa=bound,original_control_residual_MPa=control,steel_strain_utilization=steel_util,branches=branches))
    length=TestNorm(full-accepted);field_error=TestNorm(full-(2*half-accepted))/length
    lambda_error=float(abs(2*half_lam-reference_lam-full_lam));radius=float(np.sqrt(length**2+alpha**2*(full_lam-reference_lam)**2))
    joint=float(np.sqrt((field_error*length)**2+(alpha*lambda_error)**2)/radius)
    assert abs(length-native_length)<2e-10*length and abs(field_error-native_field_error)<2e-10 and abs(lambda_error-native_lambda_error)<1e-12 and abs(joint-native_joint)<2e-10
    different=[dict(point=a['point'],side=a['side'],real_concrete=a['real_concrete'],full=a['branch'],half=b['branch'],
        full_trial_yield_MPa=a['trial_yield_MPa'],half_trial_yield_MPa=b['trial_yield_MPa']) for a,b in zip(entries[0]['branches'],entries[1]['branches']) if a['branch']!=b['branch']]
    pairs.append(dict(states=entries,length=length,field_error=field_error,lambda_error=lambda_error,
        native_lambda_bound=float(native_lambda_bound),joint_error=joint,working_field_error_limit=.005,
        native_metric_independently_reproduced=True,branch_differences=different,
        max_increment_full=float(abs(np.einsum('spcd,d->spc',op,full-accepted)).max()),
        max_increment_half=float(abs(np.einsum('spcd,d->spc',op,half-accepted)).max())))
assert r.position==len(r.raw)
proof=dict(actual_primary_binary_exact947=True,passive_capture_did_not_change_outcome=True,
    native=inner['native'],all_full_and_half_equilibria_independently_verified=True,
    strain_metric_and_half_defect_independently_reproduced=True,pairs=pairs,
    source_history25_unchanged=True,all_guard_tolerances_preserved=True,
    physical_capacity_certified=False,no_new_accepted_history=True)
(root/'WorkingGuidePairProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
for p in pairs:
    print('PAIR FIELD ERROR',p['field_error'],'LAMBDA ERROR',p['lambda_error'],'BRANCH DIFFERENCES',len(p['branch_differences']))
    for state in p['states']:print(state['name'],'LAMBDA',state['lambda_value'],'R',state['independent_total_residual_MPa'])
    for change in p['branch_differences']:print('BRANCH',change)
