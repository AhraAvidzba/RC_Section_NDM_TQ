"""ДЛЯ ТЕСТОВ. Проверяет настоящий холодный запуск общего ядра, сохранённую книгу и физические масштабы без отождествления NumFail с разрушением."""
from pathlib import Path
from runpy import run_path
from zipfile import ZipFile
from argparse import ArgumentParser
import json,re,numpy as np
repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,default=base/'Cold868FixedMetricCold_20261126');args=p.parse_args();root=args.root
parent=base/'Cold867FixedMetricKernel_20261126';control=base/'Cold853LoadingSecantCold_20261125'
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
for file in (parent/'src').rglob('*'):
    if file.is_file():assert file.read_bytes()==(root/file.relative_to(parent)).read_bytes()
driver=(root/'tests/modTestCoupledElasticState.bas').read_bytes();assert driver==(parent/'tests/modTestCoupledElasticState.bas').read_bytes()
scope=json.loads((root/'Preparation.json').read_bytes())['directed_scope'];assert scope['case']=='FixedMetricKernelTrueCold96' and scope['actual_imports_this_replay']==[]
solver=(root/'src/Solver/CCoupledDeformationContinuation.cls').read_text('utf-8-sig');assert solver.count('Private Const PROBE_LIMIT As Long = 96')==1
helpers=run_path(str(here/'Check-TestPersistentHEvidence.py'))
reg,_=helpers['TestNative'](root,'OperationalArcRegressions');assert reg==dict(passed=164,failed=0,macros=5,timeout600=False,own_process_closed=True)
text=(root/'Native/HBackendStock.txt').read_text('utf-8-sig');timeout=re.search(r'WATCHDOG_FAILURE: timeout=(\d+) sec;',text)
if timeout:
    assert int(timeout[1])==600
    closure=json.loads((root/'Native/HBackendStock.txt.ProcessClosure.json').read_bytes())
    assert closure['Stage']=='TIMEOUT' and closure['Identity']['Valid'] and not closure['Passed']
    assert closure['Before']['State']=='Live' and closure['After']['State']=='Absent'
    assert closure['Identity']['PID']==closure['Before']['PID']==closure['After']['PID']
    assert str(closure['Identity']['Ticks'])==closure['Before']['StartTicksUTC']==closure['After']['StartTicksUTC']
    assert closure['ExactOwnExcelStopRequestedOnlyByExplicitTimeout'] and closure['ExcelStopError'] is None
    outcome=dict(timeout_seconds=600,workflow_result_missing=True,own_process_closed=True)
else:
    native,_=helpers['TestNative'](root,'HBackendStock');assert native['failed']==0 and native['own_process_closed']
    saved,=re.findall(r'(?m)^RESULTS_SAVE_REOPEN: .*?equal=True; file=([^\r\n]+)',text)
    saved=Path(saved).resolve();assert saved.is_relative_to(root.resolve()) and saved.suffix=='.xlsm'
    (root/'SavedWorkbook.xlsm').write_bytes(saved.read_bytes())
    with ZipFile(root/'docs/regression/Diagnostic.xlsm') as book:compiled=book.read('xl/vbaProject.bin')
    outcome=helpers['TestStock'](repo,root,compiled);assert outcome['all_saved_sources_exact']==256
read=run_path(str(here/'Read-TestColdRawCapture.py'));frames,raw=read['TestHistory'](root/'Native/Stock.RawHistory.bin')
old_frames,_=read['TestHistory'](control/'Native/Stock.RawHistory.bin');current=frames[-1];old=old_frames[-1]
assert raw['incomplete_tail_bytes']==0
prefix=0
for index in range(min(len(current['fields']),len(old['fields']))):
    if not all(np.array_equal(current[k][index],old[k][index]) for k in ('fields','fractions','lambdas')):break
    prefix+=1
ops=base/'Cold835GatedColdReplay_20261125/Native/EnvelopeEnergy.Accepted.Operators.csv'
assert (root/'src/Solver/CCoupledElasticState.cls').read_bytes()==(base/'Cold835GatedColdReplay_20261125/src/Solver/CCoupledElasticState.cls').read_bytes()
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](ops,60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)])
physical=[]
for count in sorted({n for n in (4,7,19,49,len(current['fields'])) if n<=len(current['fields'])}):
    strain=np.einsum('spcd,d->spc',op,current['fields'][count-1]);concrete=strain[:,:25];tensor=np.zeros((3,25,3,3))
    for i in range(3):tensor[:,:,i,i]=concrete[:,:,i]
    for k,(i,j) in enumerate(((0,1),(0,2),(1,2)),3):tensor[:,:,i,j]=tensor[:,:,j,i]=concrete[:,:,k]/2
    principal=np.linalg.eigvalsh(tensor);steel=np.einsum('spc,pc->sp',strain[:,25:],axis[25:]);long=axis[25:,2]==1
    limits=np.where(long[None,:],np.where(steel>=0,.025,.018),.02)
    physical.append(dict(state=count,lambda_value=float(current['lambdas'][count-1]),maximum_lambda_so_far=float(max(current['lambdas'][:count])),
        concrete_principal_min=float(principal.min()),concrete_principal_max=float(principal.max()),
        max_abs_long_steel_strain=float(np.max(abs(steel[:,long]))),max_abs_hoop_strain=float(np.max(abs(steel[:,~long]))),
        maximum_three_station_steel_utilization=float(np.max(abs(steel)/limits))))
proof=dict(regressions=reg,cold=outcome,raw_history=raw,accepted_states=len(current['fields']),negative853_exact_prefix_states=prefix,
    maximum_accepted_lambda=float(max(current['lambdas'])),last_accepted_lambda=float(current['lambdas'][-1]),physical_progress=physical,
    same_actual_compiled_kernel867=True,full_cold_from_zero=True,original_materials_and_force_tolerances=True,global_probe_limit96=True,
    no_warm_fields_passed_to_stock_macro=True,entire_new_cold_trajectory_stress_replay_still_required=True,
    physical_capacity_certified=False,whole_goal_completed=False,no_release_promotion=True)
(root/'FixedMetricColdProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('COLD96 STATES',len(current['fields']),'MAX LAMBDA',max(current['lambdas']),'LAST',current['lambdas'][-1],
    'OUTCOME',outcome.get('capacity',outcome.get('timeout_seconds')),'164 REGRESSION PASS;PHYSICAL LIMIT NOT CERTIFIED')
