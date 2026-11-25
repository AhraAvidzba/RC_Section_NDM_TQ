"""ДЛЯ ТЕСТОВ. Прослеживает проверенную холодную историю19 существующей дуговой операцией без новых исходников или будущих полей."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
source=base/'Cold857SpatialDeclineRecovery_20261125';verified=base/'Cold859SpatialDeclineStressReplay_20261125'
compiled=base/'Cold852AutomaticLoadingSecantWalk_20261125';root=base/'Cold861Cold19AutomaticArcRestart_20261125'
stress=json.loads((verified/'VerifiedColdStressReplayProof.json').read_bytes())
assert stress['source_states']==19 and stress['source_workflow_incomplete'] and stress['all_raw_fields_fractions_and_lambdas_exact']
assert stress['native']['own_process_closed']
read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'];history=read(verified/'Input/Reverse1.bin');original=read(compiled/'Input/Reverse1.bin')
for key in ('calibration','options','offset','direction'):np.testing.assert_array_equal(history[key],original[key])
for name in ('src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledArcContinuation.cls','src/Solver/CCoupledElasticState.cls',
 'src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledConcreteReturn.cls','src/Materials/CCoupledReinforcedResponse.cls'):
    assert (compiled/name).read_bytes()==(source/name).read_bytes(),name
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,compiled,1e-6,1,1e-6,0)
prep=json.loads((root/'Preparation.json').read_bytes());values={
 'Input/Reverse1.bin':(verified/'Input/Reverse1.bin').read_bytes(),
 'Input/SourceColdRawHistory.bin':(source/'Native/Stock.RawHistory.bin').read_bytes(),
 'SourceColdStressProof.json':(verified/'VerifiedColdStressReplayProof.json').read_bytes(),
 'SourceColdOutcome.json':(source/'SpatialDeclineColdProof.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
compilation=json.loads((compiled/'docs/regression/CodeUpdate.json').read_bytes())
prep['directed_scope']=dict(case='AutomaticArcRestartFromVerifiedCold19',actual_import_this_pass=compilation['Modules'],
 actual_imports_this_replay=[],actual_component_count=256,immutable_compiled_book_reused=True,compiled_source_root=compiled.name,
 source_root=source.name,verified_stress_root=verified.name,source_states=19,source_workflow_incomplete=True,
 all_arc_Newton_material_and_kinematic_methods_exact_source=True,controller_Trace_not_called=True,
 no_future_field_or_material_face_input=True,original96calls32attempts_four_per_revision=True,
 original_force_half_lambda_locality_and_steel_criteria=True,no_new_cold_capacity_claim=True,
 physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(compiled/'docs/regression/Diagnostic.xlsm').read_bytes()
print(root,'EXACT VERIFIED COLD19;IMMUTABLE COMPILED852;ARC RESTART NOT NEW COLD CAPACITY;96CALLS')
