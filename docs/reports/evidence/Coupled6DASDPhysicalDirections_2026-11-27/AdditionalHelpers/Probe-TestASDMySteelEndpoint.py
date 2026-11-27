"""ДЛЯ ТЕСТОВ. Ищет внешние равновесия My при больших деформациях стали с точным исключением линейной связи."""
from pathlib import Path
from argparse import ArgumentParser
import json
import subprocess
import sys
import traceback

parser=ArgumentParser(description=__doc__)
parser.add_argument('--root',type=Path,required=True)
parser.add_argument('--worker',action='store_true')
parser.add_argument('--refine',action='store_true')
args=parser.parse_args();root=args.root;here=Path(__file__).resolve().parent
report_name='IndependentMySteelEndpointRefined' if args.refine else 'IndependentMySteelEndpoint'
if not args.worker:
    command=[sys.executable,str(Path(__file__).resolve()),'--root',str(root),'--worker']
    if args.refine:command.append('--refine')
    process=subprocess.run(command,
                           stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,timeout=120,
                           creationflags=subprocess.CREATE_NO_WINDOW)
    assert process.returncode==0,process.returncode
    proof=json.loads((root/'Native'/f'{report_name}.json').read_bytes())
    print(json.dumps({k:v for k,v in proof.items() if k!='results'},sort_keys=True))
    for row in proof['results']:print(json.dumps({k:v for k,v in row.items() if k not in ['field','actual_resultants_N_Nmm']},sort_keys=True))
    sys.exit(0)
try:
    source=here/'Replay-TestMainCapacityEquilibrium.py'
    code=source.read_text('utf-8-sig').split('started=perf_counter()',1)[0]
    code=code.replace("source.get('three_changed_components_exact') or source.get('changed_components_exact')",
                      "source.get('all260_working_components_unchanged') and source.get('no_kernel_modifications')",1)
    code=code.replace("(root/'docs/regression/Diagnostic.xlsm')","(root.parent/'docs/regression/Diagnostic.xlsm')",1)
    data=dict(__file__=str(source),__name__='TestMySteelData');sys.argv=[str(source),'--root',str(root)]
    exec(compile(code,str(source),'exec'),data)
    assert data['path_code']==3
    native=(here/'Solve-TestASDMaterialUpper.py').read_text('utf-8-sig')
    setup=native[native.index("    ops=data['ops']"):native.index("    load=data['offset']+high*data['direction']")]
    evaluation=native[native.index('    def TestEvaluate('):native.index('    started=perf_counter()',native.index('    def TestEvaluate('))]
    body='''    case=json.loads((case_root/'Case.json').read_bytes());prefix=data['history']
'''+setup+'''
    origin=np.asarray(data['history'][-1]);origin_lambda=float(data['factors'][-1])
    previous=json.loads((case_root/'Native/IndependentMyProbes-last-trf.json').read_bytes())
    right=next(row for row in previous['results'] if row['control']=='force' and row['ratio']==1.01 and row['accepted_by_original_tolerance'])
    right_field=np.array(right['field']);right_lambda=right['lambda_value']
    q=np.mean(np.einsum('pcd,pc->pd',operators[1,concrete:concrete+2],axis[concrete:concrete+2]),axis=0)
    origin_measure=float(q@origin);right_measure=float(q@right_field);assert right_measure>origin_measure>0
    load=data['offset']+origin_lambda*data['direction'];external=data['L']@load
    force_scales=np.array([area,area*length,area*length,area*length,area,area]);concrete_E=data['contract']['concrete_E_MPa'];calls=[]
    scale=np.maximum(abs(right_field),1e-4)
'''+evaluation+'''
    results=[];started=perf_counter()
    for target in [.01,.025]:
        ratio=(target-origin_measure)/(right_measure-origin_measure)
        seed=origin+ratio*(right_field-origin);lambda_seed=origin_lambda+ratio*(right_lambda-origin_lambda)
        scale=np.maximum(abs(seed),1e-4)
        scaled_q=q*scale;anchor=target*scaled_q/(scaled_q@scaled_q);basis=null_space(scaled_q[None,:])
        assert basis.shape==(dof,dof-1)
        initial=np.r_[basis.T@(seed/scale-anchor),lambda_seed/origin_lambda]
        def TestMapped(coordinates):
            load[:]=data['offset']+(coordinates[-1]*origin_lambda)*data['direction'];external[:]=data['L']@load
            residual,field,actual,errors=TestEvaluate(anchor+basis@coordinates[:-1])
            return residual
        solved=least_squares(TestMapped,initial,jac='3-point',method='trf',x_scale='jac',diff_step=1e-5,
                             xtol=1e-12,ftol=1e-12,gtol=1e-12,max_nfev=150)
        TestMapped(solved.x);_,field,actual,errors=TestEvaluate(anchor+basis@solved.x[:-1])
        bound=float(data['absolute']+data['relative']*max(abs(external[:dof-4])))
        equivalent=np.array([ops.NDTest('GetResponse',points+point+1,'equivalentTotalStrain') for point in range(concrete)])
        strain=np.einsum('spcd,d->spc',operators,field);eps=np.einsum('spc,pc->sp',strain[:,concrete:],axis[concrete:])
        usage=float(np.max(abs(eps)/np.where(long[None,:],np.where(eps>=0,.025,.018),.02)))
        control_error=abs(concrete_E*(q@field-target));verified=max(max(errors),control_error)<=bound+1e-10
        results.append(dict(target_mean_longitudinal_steel_strain=target,lambda_value=float(load[2]/data['direction'][2]),
                            equilibrium_verified=verified,physical_residual_MPa=errors[0],gauge_residual_MPa=errors[1],
                            six_resultants_residual_MPa=errors[2],control_residual_MPa=control_error,original_bound_MPa=bound,
                            steel_utilization=usage,compression_utilization=float(equivalent[:,1].max()/case['compression_limit']),
                            solver_success=bool(solved.success),function_evaluations=int(solved.nfev),field=field.tolist(),
                            actual_resultants_N_Nmm=actual.tolist()))
        proof=dict(results=results,fixed_prefix_states=len(prefix),exact_linear_control_eliminated=True,
                   original_tolerances_unchanged=True,OpenSees_native_constitutive_updates_used=True,
                   OpenSees_structural_element_not_used=True,entire_continuous_path_not_proven=True,
                   elapsed_seconds=perf_counter()-started,physical_ultimate_confirmed=False,production_not_modified=True)
        (case_root/'Native/IndependentMySteelEndpoint.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\\n',encoding='utf-8')
    return proof
'''
    body=body.replace("case_root/'Native/IndependentMySteelEndpoint.json'",f"case_root/'Native/{report_name}.json'")
    if args.refine:
        anchor='        TestMapped(solved.x);_,field,actual,errors=TestEvaluate(anchor+basis@solved.x[:-1])'
        assert body.count(anchor)==1
        body=body.replace(anchor,'''        refinement=[]
        best=solved;best_norm=float(max(abs(TestMapped(solved.x))))
        for difference in [1e-4,1e-6,1e-3]:
            bound0=float(data['absolute']+data['relative']*max(abs(external[:dof-4])))
            if best_norm<=bound0:break
            candidate=least_squares(TestMapped,best.x,jac='3-point',method='trf',x_scale='jac',diff_step=difference,
                                    xtol=1e-13,ftol=1e-13,gtol=1e-13,max_nfev=100)
            candidate_norm=float(max(abs(TestMapped(candidate.x))))
            refinement.append(dict(difference=difference,physical_norm_MPa=candidate_norm,nfev=int(candidate.nfev)))
            if candidate_norm<best_norm:best=candidate;best_norm=candidate_norm
        solved=best
        TestMapped(solved.x);_,field,actual,errors=TestEvaluate(anchor+basis@solved.x[:-1])''',1)
        body=body.replace('solver_success=bool(solved.success),function_evaluations=int(solved.nfev),field=field.tolist(),',
                          'solver_success=bool(solved.success),function_evaluations=int(solved.nfev),refinement=refinement,field=field.tolist(),',1)
    namespace={}
    exec('import numpy as np\nimport json\nfrom scipy.optimize import least_squares\nfrom scipy.linalg import null_space\nfrom time import perf_counter\ndef TestMySteel(case_root,data):\n'+body,namespace)
    namespace['TestMySteel'](root,data);data['ops'].wipe()
except BaseException:
    (root/'Native'/f'{report_name}Failure.txt').write_text(traceback.format_exc(),encoding='utf-8')
    raise
