"""ДЛЯ ТЕСТОВ. Архивирует обратный равновесный кандидат и проверенный возврат к исходной цели без принятия вспомогательной истории."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={'749':base/'Cold749DepartureCandidateAudit_20261125','750':base/'Cold750DepartureGuideReturn_20261125'}
audit=json.loads((roots['749']/'DepartureCandidateAuditProof.json').read_bytes())
returned=json.loads((roots['750']/'DepartureGuideReturnProof.json').read_bytes())
assert audit['independent_forward_direction_check_would_also_reject']
assert returned['solved_original_target'] and returned['independent_NDS_and_original_control_verified']
assert returned['total_primary_steps']==17 and returned['accepted_revision_unchanged']>0
out=repo/'docs/reports/evidence/Coupled6DDepartureGuideReturn_2026-11-25';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('# Сырые свидетельства сохраняются побитно.\n* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
for key,root in roots.items():
    paths=[]
    for name in ('Preparation.json','CompilePoint.log','CompilePointClosure.json','PointTraceSourceComposition.json',
        'CompiledSourceReuse.json','docs/regression/CodeUpdate.json','ImportOrderAuditCorrection.json',
        'DepartureCandidateAuditProof.json','DepartureGuideReturnProof.json',
        'src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledDeformationContinuation.cls','tests/modTestCoupledElasticState.bas',
        'Input/Reverse1.bin','Input/Source745History.bin','Input/ExistingGuideCandidate.bin',
        'Stage-TestDepartureCandidateAudit.py','Stage-TestDepartureGuideReturn.py','Run.ndjson'):
        if (root/name).is_file():paths.append(root/name)
    paths += [p for p in (root/'Native').iterdir() if p.is_file() and p.suffix.lower() not in ('.xlsx','.xlsm')]
    for source in paths:
        target=out/key/source.relative_to(root);target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes(source.read_bytes())
for name in ('Stage-TestDepartureCandidateAudit.py','Repair-TestDepartureImportOrder.py','Check-TestDepartureCandidateAudit.py',
    'Stage-TestDepartureGuideReturn.py','Check-TestDepartureGuideReturn.py','Read-TestColdRawCapture.py',
    'Check-TestPersistentHEvidence.py','Check-TestActualReverseCapture.py','Validate-TestHardPrepared.py',Path(__file__).name):
    (out/name).write_bytes((here/name).read_bytes())
(out/'Summary.json').write_text(json.dumps(dict(backward_candidate749=audit,original_target_return750=returned,
    no_working_or_release_change=True,physical_capacity_certified=False,whole_goal_completed=False),ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
pins={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',departure_candidate749=audit,
    departure_original_target_return750=returned,
    previous_goal_turn_classification='progress:745earlyH rejected with realcold379.86sNumFail,747all46independent;748realall6HFD247rank61locallyconsistent.749candidatefullNDSbutbackward verified.750unacceptedreverseguide returnedoriginaltarget lambda1.05785498 in14+3=17sharedsteps; independent47state force/gauge/control checks. Notcoldautomaticcapacity; entiregoalactive.',
    next_solver_integration=dict(parent_numerical_case='Cold743UnverifiedCumulativeRecovery_20261125',
        keep_late_H_initial_policy=True,do_not_adopt_745_early_H=True,
        location='CCoupledNewtonSolver.TryDifferenceMaterialField before rejection of equilibrated auxiliary root',
        action='Try original saved projection/target from auxiliary root only as initialField, on unchanged accepted history; keep same material and final physical/force/gauge checks.',
        budget='maxIterations minus actual mIterations, combine iteration/evaluation/hybrid counters after nested solve; recursion guard, no auxiliary acceptance.',
        publication='Only original-target NDS after verification. Restore original control and expose recovered projection/uncertainty for outer unchanged direction/steel/peak checks.',
        evidence='Cold750DepartureGuideReturn_20261125: 14 auxiliary steps +3 return steps, returned lambda1.0578549815983267, independent full NDS and original control, auxiliary absent from return history.',
        required_checks='Actual255 source audit, focused regressions105, independent directed guide case, then true cold HardShear Both/Excel with same96probe30iteration budgets and saved reopen; no physical limit claim without mechanism.'))
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for p in (path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json'):p.write_bytes(raw)
print('749750 ARCHIVED',len(pins),'PINS; BACKWARD ROOT REJECTED; ORIGINAL TARGET RETURN VERIFIED; PHYSICAL LIMIT OPEN')
