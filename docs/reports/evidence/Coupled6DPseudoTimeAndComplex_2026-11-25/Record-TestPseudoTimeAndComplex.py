"""ДЛЯ ТЕСТОВ. Архивирует отрицательный резерв739 и настоящий холодный сложный контроль740, оставляя полную цель открытой."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={'739':base/'Cold739PseudoTimeProbe_20261125','740':base/'Cold740CurrentComplexStock_20261125'}
pseudo=json.loads((roots['739']/'PseudoTimeProbeProof.json').read_bytes())
complex=json.loads((roots['740']/'CurrentComplexStopProof.json').read_bytes())
assert pseudo['native']['timeout600'] and pseudo['all12_spatial_caps_and_linear_backward_errors_verified']
assert complex['native']['timeout600'] and complex['all_six_components_nonzero']
out=repo/'docs/reports/evidence/Coupled6DPseudoTimeAndComplex_2026-11-25';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('# Сырые свидетельства сохраняются побитно.\n* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
for number,root in roots.items():
    paths=[]
    for name in ['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompiledSourceReuse.json','PointTraceSourceComposition.json',
        'docs/regression/CodeUpdate.json','Run.ndjson','PseudoTimeProbeProof.json','CurrentComplexStopProof.json',
        'src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledDeformationContinuation.cls','src/Materials/CCoupledConcreteResponse.cls',
        'tests/modTestCoupledElasticState.bas','CandidateTestPseudoTimeProbe.bas.txt','Run-TestHMixedStock.ps1',
        'Input/PostCrackReplay.bin','Input/PostCrackReplayOptions.bin','Input/PostCrackSourceRawHistory.bin',
        'Stage-TestPseudoTimeProbe.py','Stage-TestCurrentComplexStock.py']:
        if (root/name).is_file():paths.append(root/name)
    paths += [p for p in (root/'Native').iterdir() if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    for source in paths:
        target=out/number/source.relative_to(root);target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes(source.read_bytes())
for name in ['Stage-TestPseudoTimeProbe.py','CandidateTestPseudoTimeProbe.bas.txt','Check-TestPseudoTimeProbe.py',
    'Stage-TestCurrentComplexStock.py','Check-TestCurrentComplexStop.py','Read-TestColdRawCapture.py',
    'Check-TestPersistentHEvidence.py','Validate-TestHardPrepared.py',Path(__file__).name]:(out/name).write_bytes((here/name).read_bytes())
(out/'Summary.json').write_text(json.dumps(dict(pseudo739=pseudo,complex740=complex,
    no_working_code_or_release_change=True,physical_capacity_certified=False,whole_goal_completed=False),indent=2,sort_keys=True)+'\n',encoding='utf-8')
pins={str(p.relative_to(out)).replace('\\','/'):sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pseudo_time739=pseudo,current_complex740=complex,
    previous_goal_turn_classification='progress:738 fullFDconfirmsrank17,checkpointeda4b7e.739bounded12nativePTCtrials/metric/backwarderrorschecked butonly3.52%normgainandnobetterload-balancedresidual; ordinaryverificationtimeout600,no NDSoutcome.740realcoldall6Bothwithhoops/grades,600timeoutat8states/nativeL1.10015517,no returnedlimit/reopen. Badprototype notpromoted;741onlymaterialclass exact711 ordinaryfirst+closurefallback staged/compiled/audited fordirectedmaterial/realcasepair. Entiregoalpreserved.')
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for p in (path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json'):p.write_bytes(raw)
print('739-740 ARCHIVED:',len(pins),'PINS; BOTH NEGATIVE OUTCOMES PRESERVED; NO PHYSICAL LIMIT CLAIM')
