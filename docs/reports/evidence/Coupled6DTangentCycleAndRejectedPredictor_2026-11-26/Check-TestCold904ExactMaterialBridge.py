"""ДЛЯ ТЕСТОВ. Заново собирает усилия904 по существующим тензорам только после точного совпадения всей траектории и материального кода."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
source=base/'Cold904ConsistentRefinementCold96_20261126';prior=base/'Cold897FreshArcCold96_20261126'
replay=base/'Cold898FreshArcColdStress_20261126';root=base/'Cold905ConsistentRefinementColdStress_20261126'
assert not root.exists()
cold=json.loads((source/'ConsistentRefinementCold96Proof.json').read_bytes())
assert cold['cold']['native']['own_process_closed'] and cold['cold']['all_saved_sources_exact']==256
previous=json.loads((replay/'VerifiedColdStressReplayProof.json').read_bytes())
assert previous['native']['own_process_closed'] and previous['source_states']==32
read=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory']
now,raw=read(source/'Native/Stock.RawHistory.bin');old,_=read(prior/'Native/Stock.RawHistory.bin');now=now[-1];old=old[-1]
assert raw['incomplete_tail_bytes']==0 and len(now['fields'])==32
binary=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary']
current_input=binary(source/'Input/Reverse1.bin');replayed=binary(replay/'Input/Reverse1.bin')
for key in ('fields','fractions','lambdas'):
    np.testing.assert_array_equal(now[key],old[key]);np.testing.assert_array_equal(now[key],replayed[key])
for key in ('calibration','options','offset','direction'):np.testing.assert_array_equal(current_input[key],replayed[key])
file=here/'Check-TestCriticalHColdStress.py';code=file.read_text('utf-8-sig')
namespace={'__file__':str(file),'__name__':'cold904_material_scope_only'}
exec(compile(code,str(file),'exec'),namespace)
namespace['source']=source;namespace['compiled']=replay
scope=namespace['TestReplayScope']()
validate=run_path(str(here/'Validate-TestHardPrepared.py'))
mechanics,fields=validate['TestValidate'](replay,'EnvelopeEnergy.Accepted')
assert len(fields)==32;np.testing.assert_allclose(fields,now['fields'],rtol=1e-14,atol=1e-300)
# LAW печатается Trim$(Str$(Double)), а OPTIONS содержит только первые шесть параметров; точный ключ уже проверен по двоичным входам выше.
reported_material=np.array([float(format(value,'.15g')) for value in replayed['calibration']])
np.testing.assert_array_equal(mechanics['material'],reported_material)
np.testing.assert_array_equal(mechanics['options'],replayed['options'][:6])
for state,lam,fraction in zip(mechanics['states'],now['lambdas'],now['fractions']):
    assert abs(state['lambda']-lam)<5e-14 and state['fraction']==fraction
    assert max(state['residual_MPa'],state['force_residual_MPa'])<=state['bound_MPa']+1e-12
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](replay/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)])
strain=np.einsum('spcd,nd->nspc',op,now['fields']);steel=np.einsum('nspc,pc->nsp',strain[:,:,25:],axis[25:])
long=axis[25:,2]==1;limits=np.where(long[None,None,:],np.where(steel>=0,.025,.018),.02)
utilization=float(np.max(abs(steel)/limits));assert utilization<=1
diagnostics=(source/'Native/Stock.Checkpoints.csv').read_text('cp1251').splitlines()
markers=[line for line in diagnostics if line.startswith('DETAILS;') and 'пространственная касательная уточнена' in line]
assert markers
root.mkdir()
pins={str(path.relative_to(base)):sha256(path.read_bytes()).hexdigest().upper() for path in (
 source/'Native/Stock.RawHistory.bin',source/'ConsistentRefinementCold96Proof.json',
 replay/'Input/Reverse1.bin',replay/'Native/EnvelopeEnergy.Accepted.csv',
 replay/'Native/EnvelopeEnergy.Accepted.Operators.csv',replay/'VerifiedColdStressReplayProof.json')}
proof=dict(source_cold=source.name,referenced_native_replay=replay.name,source_states=32,
 entire_raw_trajectory_bit_identical897_and898=True,material_and_kinematic_scope=scope,
 same_input_calibration_options_load_direction_and_offset=True,full_binary_input_keys_bit_exact=True,
 CSV_LAW_matches15digit_Str_representation=True,CSV_OPTIONS_contract_first6_parameters=True,
 CSV_material_max_textual_difference=float(np.max(abs(np.asarray(mechanics['material'])-replayed['calibration']))),
 independently_reassembled_equilibrium=mechanics,
 all_three_station_steel_limits_independently_checked=True,maximum_three_station_steel_utilization=utilization,
 refinement_seen_in_diagnostics=True,refinement_diagnostic_records=len(markers),
 cold_seconds=cold['cold']['seconds'],prior_seconds=820.4443359375,no_clean_speed_claim=True,
 no_new_native_replay_needed_for_exact_trajectory=True,no_new_NDS_or_capacity_search=True,
 native_ownership_not_fabricated=True,source_hashes=pins,physical_capacity_certified=False,whole_goal_completed=False)
(root/'VerifiedColdStressBridgeProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
(root/'Preparation.json').write_text(json.dumps(dict(case='ExactHistoryMaterialEvidenceBridge',no_new_native_process=True,
 referenced_replay=replay.name,source_cold=source.name),sort_keys=True,indent=2)+'\n',encoding='utf-8')
(root/Path(__file__).name).write_bytes(Path(__file__).read_bytes())
print('COLD904 ALL32 EXACT897/898;MATERIAL SCOPE EXACT;SIX FORCES/56+4 REBUILT;ALL3 STEEL;NO NEW NATIVE PROCESS')
print('MAX RESIDUAL',max(max(s['residual_MPa'],s['force_residual_MPa']) for s in mechanics['states']),
      'REFINEMENT RECORDS',len(markers),'NO PHYSICAL CAPACITY CLAIM')
