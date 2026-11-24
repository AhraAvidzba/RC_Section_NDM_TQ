"""ДЛЯ ТЕСТОВ. Архивирует завершённые718–722, включая отрицательную аналитику и удачную исходную кривую."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
out=repo/'docs/reports/evidence/Coupled6DPriorBranchReplay_2026-11-24'
assert not out.exists();out.mkdir(parents=True)
cases={
    '718':'Cold718ControlledLoadBalance_20261124',
    '719':'Cold719ScopedPeakRefinement_20261124',
    '720':'Cold720PostCrackRawReplay_20261124',
    '721':'Cold721PostCrackDirectProbe_20261124',
    '722':'Cold722PriorBranchClosure_20261124'}
common=['Preparation.json','CompilePointClosure.json','PointTraceSourceComposition.json','CompiledSourceReuse.json',
    'docs/regression/CodeUpdate.json','Guard-IsolatedOffice.ps1','Run.ndjson']
proofs=['ControlledBalanceScopeProof.json','PostCrackRawReplayProof.json','DirectProbeProof.json','PriorBranchClosureProof.json']
for number,name in cases.items():
    root=base/name;paths=[root/p for p in common+proofs if (root/p).is_file()]
    for folder in ('Native','Matrix/Native','Input'):
        if (root/folder).is_dir():
            paths += [p for p in (root/folder).iterdir() if p.is_file() and
                (folder!='Input' or p.name.startswith('PostCrackReplay') or p.name=='PostCrackSourceRawHistory.bin')]
    if number in ('718','719','720','722'):
        paths += [root/p for p in ('src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledDeformationContinuation.cls',
            'src/Materials/CCoupledConcreteResponse.cls','tests/modTestCoupledElasticState.bas')]
    for p in paths:
        target=out/number/p.relative_to(root);target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes(p.read_bytes())
helpers=['Stage-TestPostCrackRawReplay.py','TestPostCrackRawReplay.bas.txt','Check-TestPostCrackRawReplay.py',
    'Stage-TestPostCrackReplayProbe.py','Check-TestPostCrackDirectProbe.py','Check-TestControlledBalanceScope.py',
    'Stage-TestPriorBranchClosure.py','Check-TestPriorBranchClosure.py','Check-TestShoulderAccuracy.py',
    'Check-TestPostCrackFields.py','Validate-TestHardPrepared.py','Read-TestColdRawCapture.py',
    'Read-TestContinuationCheckpoint.py','Check-TestPersistentHEvidence.py',Path(__file__).name]
for name in helpers:(out/name).write_bytes((here/name).read_bytes())
(out/'.gitattributes').write_text('* -text whitespace=blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
pins={str(p.relative_to(out)).replace('\\','/'):sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
scope=json.loads((base/cases['719']/'ControlledBalanceScopeProof.json').read_bytes())
replay=json.loads((base/cases['720']/'PostCrackRawReplayProof.json').read_bytes())
direct=json.loads((base/cases['721']/'DirectProbeProof.json').read_bytes())
curve=json.loads((base/cases['722']/'PriorBranchClosureProof.json').read_bytes())
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
# Исправляет неоднозначный номер, не переписывая исходные нативные отчёты.
def TestQuestionNumbers(value):
    """Только ссылки на физический контракт бетона; самостоятельная тема008 СП35 сохраняется."""
    if isinstance(value,str):return value.replace('finite concrete physical-limit contract ONQ-008','finite concrete physical-limit contract ONQ-012').replace('concrete finite physical-limit contract ONQ-008','concrete finite physical-limit contract ONQ-012')
    if isinstance(value,dict):return {k:TestQuestionNumbers(v) for k,v in value.items()}
    if isinstance(value,list):return [TestQuestionNumbers(v) for v in value]
    return value
state=TestQuestionNumbers(state)
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',
    prior_branch_replay=dict(evidence=str(out.relative_to(repo)).replace('\\','/'),
        capacity718_timed_out=True,capacity718_last_states=65,
        local_peak718=scope['native_local_peak'],local_peak718_accuracy=scope['native_local_peak_accuracy'],
        regression719=scope['regression_after'],replay720=replay['native'],
        replay720_all65_equilibrium_verified=True,replay720_analytical_misses=len(replay['analytical_force_misses']),
        replay720_max_force_error_N=max(p['error_N'] for p in replay['analytical_forces']),
        direct721_solved=direct['numerical_step_solved'],direct721_error_N=direct['last_analytical']['error_N'],
        prescribed722=curve['native'],prescribed722_points=21,
        prescribed722_max_force_error_N=max(p['error_N'] for p in curve['curve']),
        physical_capacity_certified=False,working_material_unchanged=True),
    previous_goal_turn_classification='progress: exact65 native replay exposes35 analytical misses despite equilibrium; same direct step reproduces82.11N error; prior-history full-step material variant722 passes original21 curve194/0 and independent all21mechanics without loosening tolerances; automatic723 pending.')
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for p in (path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json'):p.write_bytes(raw)
print('718-722 EVIDENCE SAVED',len(pins),'FILES; NEGATIVE ANALYTICS PRESERVED; NO PHYSICAL CAPACITY CERTIFICATE')
