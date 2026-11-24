"""ДЛЯ ТЕСТОВ. Сохраняет доказательства741742 и ограничения холодного предела, не включая текущие живые проверки743."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={'741':base/'Cold741OrdinaryFirstComplex_20261125','742':base/'Cold742RealColdFrozenReplay_20261125'}
proof=json.loads((roots['741']/'OrdinaryFirstComplexProof.json').read_bytes())
replay=json.loads((roots['742']/'RealColdFrozenReplayProof.json').read_bytes())
assert proof['cold741']['capacity']=='NumFail' and replay['states']==12
out=repo/'docs/reports/evidence/Coupled6DOrdinaryFirstComplex_2026-11-25';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('# Сырые свидетельства сохраняются побитно.\n* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
for key,root in roots.items():
    paths=[]
    for name in ('Preparation.json','CompilePoint.log','CompilePointClosure.json','PointTraceSourceComposition.json',
        'CompiledSourceReuse.json','docs/regression/CodeUpdate.json','OrdinaryMaterialRegressionProof.json',
        'OrdinaryFirstComplexProof.json','RealColdFrozenReplayProof.json',
        'src/Materials/CCoupledConcreteResponse.cls','src/Solver/CCoupledNewtonSolver.cls',
        'src/Solver/CCoupledDeformationContinuation.cls','tests/modTestCoupledElasticState.bas',
        'Input/Reverse1.bin','Input/Source741History.bin','Stage-TestOrdinaryFirstComplex.py',
        'Stage-TestRealColdFrozenReplay.py','Run-TestRealColdReplay.ps1','Run.ndjson'):
        if (root/name).is_file():paths.append(root/name)
    paths += [p for p in (root/'Native').iterdir() if p.is_file() and p.suffix.lower() not in ('.xlsx','.xlsm')]
    for source in paths:
        target=out/key/source.relative_to(root);target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes(source.read_bytes())
paired=base/'Cold739PseudoTimeProbe_20261125/Native'
for source in paired.glob('PriorBranchHardShearDirect.txt*'):
    if source.is_file():target=out/'739'/source.name;target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes(source.read_bytes())
for name in ('Stage-TestOrdinaryFirstComplex.py','Check-TestOrdinaryFirstComplex.py','Stage-TestRealColdFrozenReplay.py',
    'Check-TestRealColdFrozenReplay.py','Read-TestColdRawCapture.py','Check-TestPersistentHEvidence.py',
    'Check-TestActualReverseCapture.py','Validate-TestHardPrepared.py','Run-TestPriorBranchHardShear.ps1',Path(__file__).name):
    (out/name).write_bytes((here/name).read_bytes())
(out/'Summary.json').write_text(json.dumps(dict(cold_and_pair741=proof,independent_replay742=replay,
    no_working_or_release_change=True,physical_capacity_certified=False,whole_goal_completed=False),ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
pins={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',ordinary_first_complex741=proof,
    actual_cold_frozen_replay742=replay,
    previous_goal_turn_classification='progress: completed741coldBoth131.54s DirectOK CapacityNumFail lower1.1057066 and actualsavedreopen255.742independentall12weak/gauge/all6equilibrium verified.743candidate numericalretry after cumulativepeak shoulderfailure prepared; no physicalcapacity and fullgoal unchanged.')
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for p in (path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json'):p.write_bytes(raw)
print('741742 ARCHIVED',len(pins),'PINS;12 INDEPENDENT FIELDS; CAPACITY REMAINS UNCONFIRMED')
