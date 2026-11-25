"""ДЛЯ ТЕСТОВ. Отделяет число состояний и сохранённый максимум lambda от фактического изменения деформаций и расхода бюджета."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,csv,re,numpy as np

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
warm=base/'Cold852AutomaticLoadingSecantWalk_20261125';cold=base/'Cold853LoadingSecantCold_20261125';replay=base/'Cold835GatedColdReplay_20261125'
warmproof=json.loads((warm/'AutomaticLoadingWalkProof.json').read_bytes());coldproof=json.loads((cold/'LoadingSecantColdProof.json').read_bytes())
assert warmproof['new_accepted_equilibria']==22 and warmproof['all_new_H_strictly_growing']
assert coldproof['cold832_fields_fractions_lambdas_bit_identical'] and not coldproof['independent_cold_stress_replay_pending']
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary']
source=read(warm/'Input/Reverse1.bin');fields=list(source['fields']);factors=list(source['lambdas'])
for item in warmproof['attempts']:
    if not item['solved']:continue
    r=Reader((warm/'Native'/f"LoadingWalk.{item['attempt']:03}.bin").read_bytes());packet=[r.VectorTest() for _ in range(15)]
    fields.append(r.VectorTest());factors.append(packet[0][23])
fields=np.array(fields);factors=np.array(factors);assert len(fields)==92
cold_fields=read(replay/'Input/Reverse1.bin');assert len(cold_fields['fields'])==49
operators=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators']
g,L,w,v,coords,axis,b,gradient=operators(warm/'Native/LoadingWalk.Accepted.Sequence.Operators.csv',60,43)
old_ops=operators(replay/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
assert all(np.array_equal(a,b) for a,b in zip((g,L,w,v,coords,axis,b,gradient),old_ops))
op=np.array([b+s*.001*gradient for s in (-1,0,1)])
rows=run_path(str(here/'Validate-TestHardPrepared.py'))['TestRows']

def TestPhysicalStates(root,phase,fields,factors,selected):
    """Из кинематических операторов восстанавливает тензоры, а материальную H читает из независимо проверенных напряжённых состояний."""
    energies=np.zeros((len(fields),3,25,2));seen=np.zeros((len(fields),3,25),int)
    for row in rows(root/'Native'/(phase+'.csv')):
        if row[0]=='S' and int(row[3])<=25:
            frame,side,point=int(row[1])-1,int(row[2])+1,int(row[3])-1
            energies[frame,side,point]=np.array(row[10:12],float);seen[frame,side,point]+=1
    assert np.all(seen==1)
    gt,gc,localization=source['calibration'][6:9]
    h=np.einsum('fspc,p,c->f',energies,coords[:25,2],np.array([gt/localization,gc/localization]))/(3*sum(coords[:25,2]))
    result=[]
    for count in selected:
        strain=np.einsum('spcd,d->spc',op,fields[count-1]);concrete=strain[:,:25];tensors=np.zeros((3,25,3,3))
        tensors[:,:,0,0]=concrete[:,:,0];tensors[:,:,1,1]=concrete[:,:,1];tensors[:,:,2,2]=concrete[:,:,2]
        for component,(i,j) in enumerate(((0,1),(0,2),(1,2)),3):tensors[:,:,i,j]=tensors[:,:,j,i]=concrete[:,:,component]/2
        principal=np.linalg.eigvalsh(tensors)
        steel=np.einsum('spc,pc->sp',strain[:,25:],axis[25:]);long=axis[25:,2]==1
        limit=np.where(long[None,:],np.where(steel>=0,.025,.018),.02)
        result.append(dict(state=count,lambda_value=float(factors[count-1]),maximum_lambda_so_far=float(max(factors[:count])),
            concrete_principal_min=float(np.min(principal)),concrete_principal_max=float(np.max(principal)),
            maximum_abs_long_steel_strain=float(np.max(abs(steel[:,long]))),maximum_abs_hoop_strain=float(np.max(abs(steel[:,~long]))),
            maximum_steel_strain_utilization=float(np.max(abs(steel)/limit)),H_MPa=float(h[count-1])))
    return result

physical_warm=TestPhysicalStates(warm,'LoadingWalk.Accepted.Sequence',fields,factors,(43,70,86,92))
physical_cold=TestPhysicalStates(replay,'EnvelopeEnergy.Accepted',cold_fields['fields'],cold_fields['lambdas'],(42,43,49))
records=[];header=None
for row in rows(cold/'Native/Stock.Checkpoints.csv'):
    if row[0]=='CHECKPOINT':header=list(map(int,row[1:4]))
    elif row[0]=='DETAILS':
        text=';'.join(row[1:]);probe=re.search(r'; проб: (\d+)(?:;|$)',text);arc=re.search(r'вызовов проверенной знаковой дуги (\d+); принятых направлений (\d+)',text)
        if probe and arc:records.append(dict(checkpoint=header[0],states=header[1],status=header[2],probes=int(probe[1]),arc_calls=int(arc[1]),arc_accepted=int(arc[2]),details=text))
last=records[-1];assert last['states']==49 and last['probes']==96 and last['arc_calls']==31 and last['arc_accepted']==8
assert 'восстановление True' in last['details'] and 'ошибка поля 2.84398022780062E-05' in last['details']
controller=cold/'src/Solver/CCoupledDeformationContinuation.cls';code=controller.read_text('utf-8-sig')
assert 'Private Const PROBE_LIMIT As Long = 96' in code
assert 'If mProbeCount >= PROBE_LIMIT Then\n        StopSearch "Исчерпан общий лимит проб деформационного продолжения и проверки максимума."' in code
prior=records[-2];assert prior['states']==42 and prior['probes']==73
proof=dict(warm_states=physical_warm,cold_states=physical_cold,
    additional_warm_states=6,warm_lambda_change86_to92=float(factors[91]-factors[85]),
    warm_H_change86_to92_MPa=float(physical_warm[-1]['H_MPa']-physical_warm[-2]['H_MPa']),
    cold_budget=dict(limit=96,before_last_record={k:v for k,v in prior.items() if k!='details'},
        terminal_record={k:v for k,v in last.items() if k!='details'},last_arc_restore_succeeded=True,
        stop_by_global_probe_limit=True,controller_sha256=sha256(controller.read_bytes()).hexdigest().upper()),
    cold_history_bit_identical832=True,physical_capacity_certified=False,
    increasing_state_count_does_not_prove_progress_to_physical_limit=True,whole_goal_completed=False)
(warm/'ContinuationProgressAudit.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(json.dumps(proof,ensure_ascii=False,indent=2))
