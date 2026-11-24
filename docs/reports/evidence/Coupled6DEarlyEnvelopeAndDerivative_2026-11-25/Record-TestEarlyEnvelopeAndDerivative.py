"""ДЛЯ ТЕСТОВ. Сохраняет отрицательный ранний H, независимые46 состояний и полную H-производную реальной32-точечной истории."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={str(n):base/name for n,name in ((745,'Cold745EarlyGrowingEnvelope_20261125'),
    (746,'Cold746ActualHFullDifferences_20261125'),(747,'Cold747EarlyEnvelopeFrozenReplay_20261125'),
    (748,'Cold748ActualHJacobianPermission_20261125'))}
cold=json.loads((roots['745']/'EarlyEnvelopeColdProof.json').read_bytes())
replay=json.loads((roots['747']/'EarlyEnvelopeFrozenReplayProof.json').read_bytes())
fd=json.loads((roots['748']/'ActualHFullDifferenceProof.json').read_bytes())
failed=json.loads((roots['746']/'FailedRead746Proof.json').read_bytes())
assert cold['cold745']['capacity']=='NumFail' and cold['cold745']['all_saved_sources_exact']==255
assert replay['states']==cold['accepted_states']==46 and fd['sample_count']==247
assert fd['assembled']['rank']==61 and all(not s['unavailable'] for s in fd['steps'])
out=repo/'docs/reports/evidence/Coupled6DEarlyEnvelopeAndDerivative_2026-11-25';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('# Сырые свидетельства сохраняются побитно.\n* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
identity={}
for key,root in roots.items():
    paths=[]
    for name in ('Preparation.json','CompilePoint.log','CompilePointClosure.json','PointTraceSourceComposition.json',
        'CompiledSourceReuse.json','docs/regression/CodeUpdate.json','EarlyEnvelopeRegressionProof.json',
        'EarlyEnvelopeColdProof.json','EarlyEnvelopeFrozenReplayProof.json','ActualHFullDifferenceProof.json',
        'FailedRead746Proof.json','ActualColdMaterialMeasures.json',
        'Input/Reverse1.bin','Input/Source745History.bin','src/Solver/CCoupledDeformationContinuation.cls','Run.ndjson'):
        if (root/name).is_file():paths.append(root/name)
    if key in ('746','748'):paths.append(root/'tests/modTestCoupledElasticState.bas')
    paths += [p for p in (root/'Native').iterdir() if p.is_file() and p.suffix.lower() not in ('.xlsx','.xlsm')]
    for source in paths:
        target=out/key/source.relative_to(root);target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes(source.read_bytes())
    identity[key]={name:sha256((root/name).read_bytes()).hexdigest().upper() for name in
        ('src/Materials/CCoupledConcreteResponse.cls','src/Solver/CCoupledNewtonSolver.cls','tests/modTestCoupledElasticState.bas')}
(out/'744ActualColdMaterialMeasures.json').write_bytes((base/'Cold744RecoveredColdFrozenReplay_20261125/ActualColdMaterialMeasures.json').read_bytes())
(out/'UnchangedSourceIdentity.json').write_text(json.dumps(identity,indent=2,sort_keys=True)+'\n',encoding='utf-8')
for name in ('Stage-TestEarlyGrowingEnvelope.py','Check-TestEarlyEnvelopeCold.py','Stage-TestEarlyEnvelopeFrozenReplay.py',
    'Check-TestEarlyEnvelopeFrozenReplay.py','Stage-TestActualHFullDifferences.py','Stage-TestActualHJacobianPermission.py',
    'Check-TestActualHFullDifferences.py','Stage-TestTailFullDifferences.py','Check-TestTailFullDifferences.py',
    'Analyze-TestRealColdMaterials.py','Read-TestColdRawCapture.py','Check-TestPersistentHEvidence.py',
    'Check-TestActualReverseCapture.py','Validate-TestHardPrepared.py',Path(__file__).name):
    (out/name).write_bytes((here/name).read_bytes())
(out/'Summary.json').write_text(json.dumps(dict(cold745=cold,failed_read746=failed,independent_replay747=replay,
    actual_H_derivative748=fd,no_working_or_release_change=True,physical_capacity_certified=False,whole_goal_completed=False),
    ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
pins={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',early_envelope745=cold,
    failed_H_read746=failed,early_envelope_frozen747=replay,actual_H_derivative748=fd,
    actual_cold_material_measures744=json.loads((out/'744ActualColdMaterialMeasures.json').read_bytes()),
    previous_goal_turn_classification='progress:745earlyH realcold returned379.858sNumFail lower1.05220776, regression105/0 and actualsavedreopen255.747all46independentweak/gauge/all6.746queryflagfailure preserved/fixed in748; actualall6HfullFD247samples rank61,smallsteprelativeerror1.285e-5,larger0.02252. No physicalcapacity and fullscope unchanged.')
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for p in (path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json'):p.write_bytes(raw)
print('745748 ARCHIVED',len(pins),'PINS; NEGATIVE H AND247 ACTUAL DERIVATIVES; NO PHYSICAL LIMIT')
