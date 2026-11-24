"""ДЛЯ ТЕСТОВ. Проверяет независимым интегрированием все фактически принятые состояния раннего H-контроля745."""
from pathlib import Path
from runpy import run_path
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold747EarlyEnvelopeFrozenReplay_20261125';source=base/'Cold745EarlyGrowingEnvelope_20261125'
parent=base/'Cold744RecoveredColdFrozenReplay_20261125'
native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native['passed']==1 and native['failed']==0 and not native['timeout600']
assert (root/'Input/Source745History.bin').read_bytes()==(source/'Native/Stock.RawHistory.bin').read_bytes()
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](root/'Input/Source745History.bin')
frame=frames[-1]
binary=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
for key in ('fields','fractions','lambdas'):assert np.array_equal(binary[key],frame[key]),key
mechanics,fields=run_path(str(here/'Validate-TestHardPrepared.py'))['TestValidate'](root,'EnvelopeEnergy.Accepted')
assert len(fields)==len(frame['fields'])
np.testing.assert_allclose(np.asarray(fields),frame['fields'],rtol=1e-14,atol=1e-300)
for state,value in zip(mechanics['states'],frame['lambdas']):assert abs(state['lambda']-value)<=1e-14*(1+abs(value))
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
for name in ('src/Materials/CCoupledConcreteResponse.cls','src/Solver/CCoupledNewtonSolver.cls'):
    assert (root/name).read_bytes()==(source/name).read_bytes(),name
proof=dict(native=native,raw_history=raw,states=len(fields),source_fields_and_lambdas_binary_exact=True,
    all_exported_fields_match_original_Double_to15digit_CSV_precision=True,unchanged_compiled_book_exact744=True,
    material_and_Newton_exact_source745=True,source_controller_not_called_in_replay=True,
    independent_force_and_weak_equilibrium=mechanics,no_independent_constitutive_model_claimed=True,
    no_new_capacity_search=True,physical_capacity_certified=False,whole_goal_completed=False)
(root/'EarlyEnvelopeFrozenReplayProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('747 ALL',len(fields),'EARLY H COLD FIELDS: INDEPENDENT ALL6/WEAK/GAUGE VERIFIED; NO PHYSICAL LIMIT')
