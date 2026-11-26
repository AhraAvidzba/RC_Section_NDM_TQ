"""ДЛЯ ТЕСТОВ. После полного побитного совпадения923/920 переносит проверку25 принятых состояний по существующим тензорам921."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
source=base/'Cold923RecordedNewtonCold96_20261126';prior=base/'Cold920AffineRecoveryCold96_20261126'
replay=base/'Cold921AffineRecoveryColdStress_20261126';root=base/'Cold924RecordedHistoryStressBridge_20261126'
assert not root.exists()
cold=json.loads((source/'RecordedNewtonCold96Proof.json').read_bytes())
assert cold['entire_raw_trajectory_and_intermediate_frames_bit_identical920'] and cold['all257_saved_sources_verified']
assert (source/'Native/Stock.RawHistory.bin').read_bytes()==(prior/'Native/Stock.RawHistory.bin').read_bytes()
verified=json.loads((replay/'VerifiedColdStressReplayProof.json').read_bytes())
assert verified['source_states']==25 and verified['all_raw_fields_fractions_and_lambdas_exact']
assert verified['native']['own_process_closed'] and verified['all_six_forces_and_all_56_physical_and4_gauge_equations']
assert verified['all_three_station_steel_limits_checked']
file=here/'Check-TestCriticalHColdStress.py'
scope={'__file__':str(file),'__name__':'recorded_cold_material_scope_test'}
exec(compile(file.read_text('utf-8-sig'),str(file),'exec'),scope)
scope['source']=source;scope['compiled']=replay
material_scope=scope['TestReplayScope']()
read=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'];frames,_=read(source/'Native/Stock.RawHistory.bin');now=frames[-1]
binary=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary']
key=binary(source/'Input/Reverse1.bin');replayed=binary(replay/'Input/Reverse1.bin')
for name in ('calibration','options','offset','direction'):assert key[name].tobytes()==replayed[name].tobytes(),name
for name in ('fields','fractions','lambdas'):assert now[name].tobytes()==replayed[name].tobytes(),name
validate=run_path(str(here/'Validate-TestHardPrepared.py'));mechanics,fields=validate['TestValidate'](replay,'EnvelopeEnergy.Accepted')
assert len(fields)==25;np.testing.assert_allclose(fields,now['fields'],rtol=1e-14,atol=1e-300)
np.testing.assert_array_equal(mechanics['material'],np.array([float(format(x,'.15g')) for x in key['calibration']]))
np.testing.assert_array_equal(mechanics['options'],key['options'][:6])
for item in mechanics['states']:assert max(item['residual_MPa'],item['force_residual_MPa'])<=item['bound_MPa']+1e-12
root.mkdir()
files=[source/'Native/Stock.RawHistory.bin',source/'RecordedNewtonCold96Proof.json',replay/'Input/Reverse1.bin',
    replay/'Native/EnvelopeEnergy.Accepted.csv',replay/'Native/EnvelopeEnergy.Accepted.Operators.csv',replay/'VerifiedColdStressReplayProof.json']
pins={p.relative_to(base).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in files}
proof=dict(source_cold=source.name,reference_cold=prior.name,referenced_native_replay=replay.name,source_states=25,
    entire_raw_history_and_intermediate_frames_bit_identical=True,material_and_kinematic_scope=material_scope,
    full_binary_input_keys_bit_exact=True,all_final25_fields_fractions_lambdas_bit_exact=True,
    independently_reassembled_equilibrium=mechanics,all_three_station_steel_limits_checked=True,
    maximum_three_station_steel_utilization=verified['maximum_three_station_steel_utilization'],
    no_new_native_process_or_material_evaluation=True,no_new_NDS_or_capacity_search=True,
    intermediate_removed_history_prefixes_not_independently_replayed_by_this_bridge=True,
    source_hashes=pins,physical_capacity_certified=False,whole_goal_completed=False)
(root/'VerifiedRecordedStressBridgeProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
(root/'Preparation.json').write_text(json.dumps(dict(case='RecordedColdExactHistoryStressBridge',no_new_native_process=True),sort_keys=True,indent=2)+'\n',encoding='utf-8')
(root/Path(__file__).name).write_bytes(Path(__file__).read_bytes())
print('ALL25 RECORDED COLD STATES REASSEMBLED FROM EXACT921 TENSORS;NO DUPLICATE NATIVE REPLAY')
