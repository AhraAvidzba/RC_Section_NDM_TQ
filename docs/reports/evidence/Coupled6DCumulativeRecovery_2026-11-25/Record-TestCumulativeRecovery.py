"""ДЛЯ ТЕСТОВ. Архивирует контроллер743, фактический холодный исход и независимый повтор744; физический предел не объявляет."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={'743':base/'Cold743UnverifiedCumulativeRecovery_20261125','744':base/'Cold744RecoveredColdFrozenReplay_20261125'}
cold=json.loads((roots['743']/'CumulativeRecoveryColdProof.json').read_bytes())
replay=json.loads((roots['744']/'RecoveredColdFrozenReplayProof.json').read_bytes())
assert cold['regression']['passed']==105 and cold['regression']['failed']==0
assert cold['cold743']['native']['own_process_closed'] and replay['native']['own_process_closed']
assert replay['states']==cold['accepted_states']
out=repo/'docs/reports/evidence/Coupled6DCumulativeRecovery_2026-11-25';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('# Сырые свидетельства сохраняются побитно.\n* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
identity={}
for key,root in roots.items():
    paths=[]
    for name in ('Preparation.json','CompilePoint.log','CompilePointClosure.json','PointTraceSourceComposition.json',
        'CompiledSourceReuse.json','docs/regression/CodeUpdate.json','CumulativeRecoveryRegressionProof.json',
        'CumulativeRecoveryColdProof.json','RecoveredColdFrozenReplayProof.json',
        'src/Solver/CCoupledDeformationContinuation.cls','Input/Reverse1.bin','Input/Source743History.bin',
        'Stage-TestUnverifiedCumulativeRecovery.py','Stage-TestRecoveredColdFrozenReplay.py',
        'Run-TestRealColdReplay.ps1','Run-TestWorkingGuideRegressions.ps1','Run.ndjson'):
        if (root/name).is_file():paths.append(root/name)
    paths += [p for p in (root/'Native').iterdir() if p.is_file() and p.suffix.lower() not in ('.xlsx','.xlsm')]
    for source in paths:
        target=out/key/source.relative_to(root);target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes(source.read_bytes())
    identity[key]={name:sha256((root/name).read_bytes()).hexdigest().upper() for name in
        ('src/Materials/CCoupledConcreteResponse.cls','src/Solver/CCoupledNewtonSolver.cls','tests/modTestCoupledElasticState.bas')}
material=(base/'Cold742RealColdFrozenReplay_20261125/ActualColdMaterialMeasures.json').read_bytes()
(out/'742ActualColdMaterialMeasures.json').write_bytes(material)
(out/'UnchangedSourceIdentity.json').write_text(json.dumps(identity,indent=2,sort_keys=True)+'\n',encoding='utf-8')
for name in ('Stage-TestUnverifiedCumulativeRecovery.py','Check-TestCumulativeRecoveryCold.py',
    'Stage-TestRecoveredColdFrozenReplay.py','Check-TestRecoveredColdFrozenReplay.py','Analyze-TestRealColdMaterials.py',
    'Read-TestColdRawCapture.py','Check-TestPersistentHEvidence.py','Check-TestActualReverseCapture.py',
    'Validate-TestHardPrepared.py',Path(__file__).name):
    (out/name).write_bytes((here/name).read_bytes())
(out/'Summary.json').write_text(json.dumps(dict(cold743=cold,independent_replay744=replay,
    no_working_or_release_change=True,physical_capacity_certified=False,whole_goal_completed=False),ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
pins={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',cumulative_recovery743=cold,
    recovered_cold_frozen_replay744=replay,actual_cold_material_measures742=json.loads(material),
    previous_goal_turn_classification='progress:741coldBothreturned131.54sNumFail and savedreopen255;742all12independentweak/gauge/all6;743cumulative numericaldeclineretry implemented privately with105/0 and actualcold evidence;744allnewstates independently integrated. No physicalcapacity and entiregoal active.')
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for p in (path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json'):p.write_bytes(raw)
print('743744 ARCHIVED',len(pins),'PINS;',cold['accepted_states'],'INDEPENDENT FIELDS; NO PHYSICAL LIMIT')
