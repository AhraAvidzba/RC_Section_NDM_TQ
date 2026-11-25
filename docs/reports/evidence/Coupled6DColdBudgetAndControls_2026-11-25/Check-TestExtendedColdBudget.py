"""ДЛЯ ТЕСТОВ. Проверяет холодный эксперимент192 и сравнивает деформации с неизменённым контролем96."""
from pathlib import Path
from runpy import run_path
from zipfile import ZipFile
from hashlib import sha256
import json,re,numpy as np
from argparse import ArgumentParser

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__);group=p.add_mutually_exclusive_group();group.add_argument('--balanced-initial',action='store_true');group.add_argument('--spatial-decline',action='store_true');args=p.parse_args()
root=base/('Cold857SpatialDeclineRecovery_20261125' if args.spatial_decline else ('Cold855BalancedInitialControl_20261125' if args.balanced_initial else 'Cold854ExtendedColdBudget192_20261125'))
parent=base/'Cold851LoadingSecantReparameterization_20261125';negative=base/'Cold853LoadingSecantCold_20261125'
name='src/Solver/CCoupledDeformationContinuation.cls';code=(root/name).read_bytes()
if args.spatial_decline:
    routing=json.loads((root/'SpatialDeclineRoutingProof.json').read_bytes())
    assert routing['only_three_Trace_routing_conditions_changed'] and routing['existing_ReplayUnverifiedDecline_exact']
    assert code.count(b'Private Const PROBE_LIMIT As Long = 96')==1
elif args.balanced_initial:
    selector=json.loads((root/'BalancedInitialControlProof.json').read_bytes())
    assert selector['only_Trace_initial_selection_call_changed'] and selector['all_other_working_methods_and_materials_exact851']
    assert code.count(b'Private Const PROBE_LIMIT As Long = 96')==1
else:
    assert code.count(b'Private Const PROBE_LIMIT As Long = 192')==1
    assert code.replace(b'Private Const PROBE_LIMIT As Long = 192',b'Private Const PROBE_LIMIT As Long = 96',1)==(parent/name).read_bytes()
    assert (root/'tests/modTestCoupledElasticState.bas').read_bytes()==(parent/'tests/modTestCoupledElasticState.bas').read_bytes()
for name in ('src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledArcContinuation.cls',
    'src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledReinforcedResponse.cls','src/Materials/CCoupledConcreteReturn.cls'):
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
text=(root/'Native/HBackendStock.txt').read_text('utf-8-sig')
timeout=re.search(r'WATCHDOG_FAILURE: timeout=(\d+) sec;',text)
if timeout:
    assert int(timeout[1])==(600 if (args.balanced_initial or args.spatial_decline) else 1200)
    closure=json.loads((root/'Native/HBackendStock.txt.ProcessClosure.json').read_bytes())
    assert closure['Stage']=='TIMEOUT' and not closure['Passed'] and closure['Identity']['Valid']
    assert closure['Before']['State']=='Live' and closure['After']['State']=='Absent'
    assert closure['Identity']['PID']==closure['Before']['PID']==closure['After']['PID']
    assert str(closure['Identity']['Ticks'])==closure['Before']['StartTicksUTC']==closure['After']['StartTicksUTC']
    assert closure['ExactOwnExcelStopRequestedOnlyByExplicitTimeout'] and closure['ExcelStopError'] is None
    outcome=dict(timeout_seconds=int(timeout[1]),workflow_result_missing=True,own_process_closed=True)
else:
    assert 'WATCHDOG_FAILURE:' not in text
    helpers=run_path(str(here/'Check-TestPersistentHEvidence.py'));native,_=helpers['TestNative'](root,'HBackendStock')
    assert native['failed']==0 and native['own_process_closed'] and not native['timeout600']
    saved,=re.findall(r'(?m)^RESULTS_SAVE_REOPEN: .*?equal=True; file=([^\r\n]+)',text)
    saved=Path(saved).resolve();assert saved.is_relative_to(root.resolve()) and saved.suffix=='.xlsm'
    (root/'SavedWorkbook.xlsm').write_bytes(saved.read_bytes())
    with ZipFile(root/'docs/regression/Diagnostic.xlsm') as book:compiled=book.read('xl/vbaProject.bin')
    outcome=helpers['TestStock'](repo,root,compiled);assert outcome['all_saved_sources_exact']==256
read=run_path(str(here/'Read-TestColdRawCapture.py'));frames,raw=read['TestHistory'](root/'Native/Stock.RawHistory.bin')
oldframes,_=read['TestHistory'](negative/'Native/Stock.RawHistory.bin');current=frames[-1];old=oldframes[-1]
prefix=0
for i in range(min(len(current['fields']),len(old['fields']))):
    if not all(np.array_equal(current[k][i],old[k][i]) for k in ('fields','fractions','lambdas')):break
    prefix+=1
assert prefix>=4 and raw['incomplete_tail_bytes']==0
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](base/'Cold835GatedColdReplay_20261125/Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
assert (root/'src/Solver/CCoupledElasticState.cls').read_bytes()==(base/'Cold835GatedColdReplay_20261125/src/Solver/CCoupledElasticState.cls').read_bytes()
op=np.array([B+s*.001*gradient for s in (-1,0,1)])
physical=[]
for count in sorted({n for n in (4,43,49,70,92,len(current['fields'])) if n<=len(current['fields'])}):
    strain=np.einsum('spcd,d->spc',op,current['fields'][count-1]);concrete=strain[:,:25];tensor=np.zeros((3,25,3,3))
    for i in range(3):tensor[:,:,i,i]=concrete[:,:,i]
    for k,(i,j) in enumerate(((0,1),(0,2),(1,2)),3):tensor[:,:,i,j]=tensor[:,:,j,i]=concrete[:,:,k]/2
    principal=np.linalg.eigvalsh(tensor);steel=np.einsum('spc,pc->sp',strain[:,25:],axis[25:]);long=axis[25:,2]==1
    limits=np.where(long[None,:],np.where(steel>=0,.025,.018),.02)
    physical.append(dict(state=count,lambda_value=float(current['lambdas'][count-1]),maximum_lambda_so_far=float(max(current['lambdas'][:count])),
        concrete_principal_min=float(principal.min()),concrete_principal_max=float(principal.max()),
        max_abs_long_steel_strain=float(np.max(abs(steel[:,long]))),max_abs_hoop_strain=float(np.max(abs(steel[:,~long]))),
        maximum_three_station_steel_utilization=float(np.max(abs(steel)/limits))))
proof=dict(cold=outcome,raw_history=raw,accepted_states=len(current['fields']),negative96_exact_prefix_states=prefix,
    maximum_accepted_lambda=float(max(current['lambdas'])),last_accepted_lambda=float(current['lambdas'][-1]),physical_progress=physical,
    single_controller_constant96_to192_only=not (args.balanced_initial or args.spatial_decline),balanced_initial_control=args.balanced_initial or args.spatial_decline,
    spatial_generalized_decline_recovery=args.spatial_decline,
    original_materials_and_tolerances_unchanged=True,
    old96_control_retained=True,default_working_budget_not_changed=True,full_cold_from_zero=True,
    independent_new_stress_replay_pending=len(current['fields'])>prefix,physical_capacity_certified=False,whole_goal_completed=False)
(root/('SpatialDeclineColdProof.json' if args.spatial_decline else ('BalancedInitialColdProof.json' if args.balanced_initial else 'ExtendedColdBudgetProof.json'))).write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('SPATIAL DECLINE96 COLD' if args.spatial_decline else ('BALANCED INITIAL96 COLD' if args.balanced_initial else 'RESEARCH192 COLD'),outcome.get('capacity','TIMEOUT'),'STATES',len(current['fields']),'MAX',proof['maximum_accepted_lambda'],'EXACT96PREFIX',prefix)
for item in physical:print(item)
