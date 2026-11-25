"""ДЛЯ ТЕСТОВ. Архивирует рабочую дуговую операцию, контекст отказа и независимый контроль60 холодных состояний без заявления предела."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from datetime import datetime
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={k:base/n for k,n in {'785':'Cold785OperationalArcController_20261125','786':'Cold786ArcFailureContext_20261125',
    '787':'Cold787ArcContextOrder_20261125','788':'Cold788OperationalColdReplay_20261125'}.items()}
proofs={k:json.loads((roots[k]/n).read_bytes()) for k,n in {'785':'OperationalColdProof.json','787':'ArcFailureContextProof.json','788':'OperationalColdReplayProof.json'}.items()}
assert proofs['785']['accepted_states']==60 and proofs['785']['cold']['all_saved_sources_exact']==256
assert proofs['785']['arc_recovered_directions']==0
assert proofs['787']['exact_context_selected'] and proofs['787']['foreign_projection_rejected'] and proofs['787']['foreign_revision_rejected']
assert proofs['788']['source_cold785_exact_all60']
native_check=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative']
for key in ('785','787','788'):
    native,_=native_check(roots[key],'RealColdReplay')
    assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
for key in ('785','787'):
    regression,_=native_check(roots[key],'OperationalArcRegressions')
    assert regression==dict(passed=164,failed=0,macros=5,timeout600=False,own_process_closed=True)
    proofs['regression'+key]=regression
state_path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(state_path.read_bytes())
cold787_path=roots['787']/'OperationalColdProof.json'
pending=not cold787_path.exists()
if not pending:
    proofs['cold787']=json.loads(cold787_path.read_bytes())
    native_check(roots['787'],'HBackendStock')
out=repo/'docs/reports/evidence/Coupled6DOperationalArcController_2026-11-25'
assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('# Сырые свидетельства сохраняются побитно.\n* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
for key,root in roots.items():
    for relative in ('OperationalArcProof.json','OperationalColdProof.json','OperationalColdReplayProof.json','ArcFailureContextProof.json',
        'Preparation.json','CompilePoint.log','CompilePointClosure.json','PointTraceSourceComposition.json','CompiledSourceReuse.json',
        'docs/regression/CodeUpdate.json','ParentCodeUpdate.json','ParentCompileClosure.json',
        'src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledArcContinuation.cls','src/Solver/CCoupledDeformationContinuation.cls',
        'src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledConcreteReturn.cls','tests/modTestCoupledElasticState.bas',
        'Compile-TestExternalPoint.ps1','Run-TestRealColdReplay.ps1','Run-TestHMixedStock.ps1','Guard-IsolatedOffice.ps1',
        'tools/build_workbook/Update-IsolatedOptimizerCode.ps1'):
        p=root/relative
        if p.is_file():
            t=out/key/relative;t.parent.mkdir(parents=True,exist_ok=True);t.write_bytes(p.read_bytes())
    for directory in ('Input','Native'):
        for p in (root/directory).iterdir():
            if not p.is_file() or p.suffix.lower() in ('.xlsm','.xlsx'):continue
            if key=='787' and pending and (p.name.startswith(('Stock.','HBackendStock.'))):continue
            t=out/key/directory/p.name;t.parent.mkdir(parents=True,exist_ok=True);t.write_bytes(p.read_bytes())
helpers=('CandidateArcContinuation.cls.txt','CandidateArcController.cls.txt','Stage-TestOperationalArc.py','Finalize-TestOperationalArc.py',
    'Check-TestOperationalArc.py','Stage-TestArcFailureContext.py','Stage-TestArcContextOrder.py','Run-TestOperationalArcRegressions.ps1',
    'Check-TestOperationalCold.py','Stage-TestOperationalColdReplay.py','Check-TestOperationalColdReplay.py',
    'Stage-TestEnergyNumerics.py','Stage-TestColdRawReplay.py','TestArcYieldBoundary.cls.txt','Compile-TestExternalPoint.ps1',
    'Finalize-TestStockYield.py','Audit-TestPrivateCoupledBook.py','Inspect-TestResolutionVBA.py','Canonical-TestCoreVBA.py',
    'Read-TestColdRawCapture.py','Check-TestActualReverseCapture.py','Check-TestPersistentHEvidence.py','Validate-TestHardPrepared.py',
    'Check-TestHardHalfSteps.py',Path(__file__).name)
for name in helpers:(out/name).write_bytes((here/name).read_bytes())
release={n:sha256((repo/n).read_bytes()).hexdigest().upper() for n in ('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt')}
assert release['workbook/output/RC_Section_NDM.xlsm']=='E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
assert release['workbook/output/VBA_All_Code.txt']=='DB8023F9384B0059E1D1C51D59FE0E62F9B8CD94AC4D9B329DD0EE810AF676CC'
summary=dict(proofs=proofs,release_hashes=release,cold787_pending=pending,
    compiled786_test_harness_not_run_before_order_correction787=True,
    sources_and_tests_private_only=True,physical_capacity_certified=False,whole_goal_completed=False)
(out/'Summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
pins={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',
    operational_arc_controller_evidence='docs/reports/evidence/Coupled6DOperationalArcController_2026-11-25',
    operational_cold785_summary={k:proofs['785'][k] for k in ('accepted_states','maximum_accepted_lambda','arc_calls','arc_recovered_directions')},
    previous_goal_turn_classification='progress: implemented operational arc class and stalled-controller integration; actual256 components audited; warm operation5calls/7primary and34 equilibria independently verified. Cold785 completed60 states, saved/reopened256 sources, NumFail retained;788 independently verified all60. Found original failure dispatch lost after auxiliary successful probe; fixed exact revision/projection routing786/787, native match/foreign-context tests passed;164 affected regression checks passed for785 and787. Physical capacity not certified.',
    next_action='Revalidate and poll existing cold787 handle65581 (not restart). On terminal outcome verify saved workbook and independently replay accepted cold history. Arc was not admitted in cold785 (4calls/0directions), so do not claim cold arc success. To resolve remaining issue capture/reproduce actual rejected arc from exact cold prefix10 (ratios1 and0.5), inspect event-half error and accepted-root residual/correction before adding retries or reducing steps.41/60 cold785 states satisfy original but not10000x tighter inner tolerance; causal role unproved. Check bounded adaptive retry at same accepted history and preserve original failure context/budgets, then physical criteria, mesh/step/external controls and full original scope. Production promotion remains unapproved/unverified.')
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for p in (state_path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json'):p.write_bytes(raw)
print('785-788 ARCHIVED',len(pins),'PINS; COLD787_PENDING',pending,'GOAL ACTIVE; NO PHYSICAL CAPACITY')
