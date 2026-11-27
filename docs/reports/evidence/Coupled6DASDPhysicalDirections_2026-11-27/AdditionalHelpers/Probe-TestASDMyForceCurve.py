"""ДЛЯ ТЕСТОВ. Продолжает внешние равновесия My от сохранённой допустимой истории, проверяя каждое до принятия."""
from pathlib import Path
from argparse import ArgumentParser
import json
import subprocess
import sys
import traceback

parser=ArgumentParser(description=__doc__)
parser.add_argument('--root',type=Path,required=True)
parser.add_argument('--worker',action='store_true')
parser.add_argument('--verified-seed',action='store_true')
args=parser.parse_args();root=args.root;here=Path(__file__).resolve().parent
report_name='IndependentMyForceCurveVerifiedSeed' if args.verified_seed else 'IndependentMyForceCurve'
if not args.worker:
    command=[sys.executable,str(Path(__file__).resolve()),'--root',str(root),'--worker']
    if args.verified_seed:command.append('--verified-seed')
    process=subprocess.run(command,
                           stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,timeout=150,
                           creationflags=subprocess.CREATE_NO_WINDOW)
    assert process.returncode==0,process.returncode
    proof=json.loads((root/'Native'/f'{report_name}.json').read_bytes())
    print(json.dumps({k:v for k,v in proof.items() if k!='records'},sort_keys=True))
    sys.exit(0)
try:
    source=here/'Replay-TestMainCapacityEquilibrium.py'
    code=source.read_text('utf-8-sig').split('started=perf_counter()',1)[0]
    code=code.replace("source.get('three_changed_components_exact') or source.get('changed_components_exact')",
                      "source.get('all260_working_components_unchanged') and source.get('no_kernel_modifications')",1)
    code=code.replace("(root/'docs/regression/Diagnostic.xlsm')","(root.parent/'docs/regression/Diagnostic.xlsm')",1)
    data=dict(__file__=str(source),__name__='TestMyCurveData');sys.argv=[str(source),'--root',str(root)]
    exec(compile(code,str(source),'exec'),data)
    assert data['path_code']==3
    native=(here/'Solve-TestASDMaterialUpper.py').read_text('utf-8-sig')
    setup=native[native.index("    ops=data['ops']"):native.index("    load=data['offset']+high*data['direction']")]
    evaluation=native[native.index('    def TestEvaluate('):native.index('    started=perf_counter()',native.index('    def TestEvaluate('))]
    body='''    case=json.loads((case_root/'Case.json').read_bytes());prefix=data['history']
'''+setup+'''
    field=np.asarray(data['history'][-1]);current=float(data['factors'][-1]);origin=current;step=.01*origin
    load=data['offset']+current*data['direction'];external=data['L']@load
    force_scales=np.array([area,area*length,area*length,area*length,area,area]);concrete_E=data['contract']['concrete_E_MPa'];calls=[]
    scale=np.maximum(abs(field),1e-4)
'''+evaluation+'''
    records=[];accepted=0;highest=current;started=perf_counter();previous=None;previous_lambda=None
    for index in range(12):
        target=current+step;load[:]=data['offset']+target*data['direction'];external[:]=data['L']@load
        seed=field*(target/current) if previous is None else field+(field-previous)*((target-current)/(current-previous_lambda))
        scale=np.maximum(abs(seed),1e-4)
        solved=least_squares(lambda q:TestEvaluate(q)[0],seed/scale,jac='3-point',method='trf',x_scale='jac',diff_step=1e-5,
                             xtol=1e-12,ftol=1e-12,gtol=1e-12,max_nfev=100)
        _,candidate,actual,errors=TestEvaluate(solved.x)
        bound=float(data['absolute']+data['relative']*max(abs(external[:dof-4])))
        equivalent=np.array([ops.NDTest('GetResponse',points+point+1,'equivalentTotalStrain') for point in range(concrete)])
        strain=np.einsum('spcd,d->spc',operators,candidate);eps=np.einsum('spc,pc->sp',strain[:,concrete:],axis[concrete:])
        steel_usage=float(np.max(abs(eps)/np.where(long[None,:],np.where(eps>=0,.025,.018),.02)))
        compression_usage=float(equivalent[:,1].max()/case['compression_limit']);equilibrium=max(errors)<=bound+1e-10
        admissible=equilibrium and steel_usage<=1 and compression_usage<=1
        records.append(dict(index=index+1,lambda_value=target,accepted=admissible,equilibrium_verified=equilibrium,
                            physical_residual_MPa=errors[0],gauge_residual_MPa=errors[1],six_resultants_residual_MPa=errors[2],
                            original_bound_MPa=bound,compression_utilization=compression_usage,steel_utilization=steel_usage,
                            solver_success=bool(solved.success),function_evaluations=int(solved.nfev),field=candidate.tolist(),
                            actual_resultants_N_Nmm=actual.tolist()))
        if admissible:
            previous=field.copy();previous_lambda=current;field=candidate;current=target;accepted+=1;highest=max(highest,current)
            TestConcrete(strain,True);steel,delta=TestSteel(strain);plastic+=delta;back+=H*delta
        proof=dict(records=records,initial_fixed_prefix_states=len(prefix),accepted_additional_states=accepted,
                   highest_admissible_lambda=highest,elapsed_seconds=perf_counter()-started,
                   original_tolerances_unchanged=True,OpenSees_native_constitutive_updates_used=True,
                   OpenSees_structural_element_not_used=True,production_not_modified=True,
                   physical_ultimate_confirmed=False,whole_goal_completed=False)
        (case_root/'Native/IndependentMyForceCurve.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\\n',encoding='utf-8')
        if equilibrium and not admissible:break
        if not equilibrium:
            step*=.5
            if step<origin*1e-5:break
    return proof
'''
    body=body.replace("case_root/'Native/IndependentMyForceCurve.json'",f"case_root/'Native/{report_name}.json'")
    if args.verified_seed:
        anchor="        scale=np.maximum(abs(seed),1e-4)"
        assert body.count(anchor)==1
        body=body.replace(anchor,'''        if index==0:
            reference=json.loads((case_root/'Native/IndependentMyProbes-last-trf.json').read_bytes())
            known=next(row for row in reference['results'] if row['control']=='force' and row['ratio']==1.01 and row['accepted_by_original_tolerance'])
            assert abs(target-known['lambda_value'])<1e-12
            seed=np.array(known['field'])
        scale=np.maximum(abs(field),1e-4)''',1)
        anchor="        solved=least_squares(lambda q:TestEvaluate(q)[0],seed/scale,jac='3-point',method='trf',x_scale='jac',diff_step=1e-5,\n                             xtol=1e-12,ftol=1e-12,gtol=1e-12,max_nfev=100)"
        assert body.count(anchor)==1
        body=body.replace(anchor,'''        initial_residual=TestEvaluate(seed/scale)[0]
        initial_bound=float(data['absolute']+data['relative']*max(abs(external[:dof-4])))
        if max(abs(initial_residual))<=initial_bound:
            solved=SimpleNamespace(x=seed/scale,success=True,nfev=1)
        else:
            solved=least_squares(lambda q:TestEvaluate(q)[0],seed/scale,jac='3-point',method='trf',x_scale='jac',diff_step=1e-5,
                                 xtol=1e-12,ftol=1e-12,gtol=1e-12,max_nfev=100)''',1)
        body=body.replace('initial_fixed_prefix_states=len(prefix),accepted_additional_states=accepted,',
                          'initial_fixed_prefix_states=len(prefix),first_seed_from_verified_global_root=True,accepted_additional_states=accepted,',1)
    namespace={}
    exec('import numpy as np\nimport json\nfrom scipy.optimize import least_squares\nfrom time import perf_counter\nfrom types import SimpleNamespace\ndef TestMyCurve(case_root,data):\n'+body,namespace)
    namespace['TestMyCurve'](root,data);data['ops'].wipe()
except BaseException:
    (root/'Native'/f'{report_name}Failure.txt').write_text(traceback.format_exc(),encoding='utf-8')
    raise
