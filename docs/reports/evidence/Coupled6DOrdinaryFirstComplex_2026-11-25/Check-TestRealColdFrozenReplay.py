"""ДЛЯ ТЕСТОВ. Независимо интегрирует все уравнения и шесть усилий двенадцати исходных холодных состояний741."""
from pathlib import Path
from runpy import run_path
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold742RealColdFrozenReplay_20261125';parent=base/'Cold741OrdinaryFirstComplex_20261125'
native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native['passed']==1 and native['failed']==0 and not native['timeout600']
assert (root/'Input/Source741History.bin').read_bytes()==(parent/'Native/Stock.RawHistory.bin').read_bytes()
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](root/'Input/Source741History.bin')
frame=frames[-1]
binary=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
for key in ('fields','fractions','lambdas'):assert np.array_equal(binary[key],frame[key]),key
phase='EnvelopeEnergy.Accepted'
mechanics,fields=run_path(str(here/'Validate-TestHardPrepared.py'))['TestValidate'](root,phase)
assert len(fields)==len(frame['fields'])==12
np.testing.assert_allclose(np.asarray(fields),frame['fields'],rtol=1e-14,atol=1e-300)
for state,value in zip(mechanics['states'],frame['lambdas']):assert abs(state['lambda']-value)<=1e-14*(1+abs(value))
for name in ('src/Materials/CCoupledConcreteResponse.cls','src/Solver/CCoupledNewtonSolver.cls',
             'src/Solver/CCoupledDeformationContinuation.cls'):
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
proof=dict(native=native,raw_history=raw,states=12,source_fields_and_lambdas_binary_exact=True,
    all_exported_fields_match_original_Double_to15digit_CSV_precision=True,
    numerical_core_exact741=True,independent_force_and_weak_equilibrium=mechanics,
    no_independent_constitutive_model_claimed=True,no_new_capacity_search=True,
    physical_capacity_certified=False,whole_goal_completed=False)
(root/'RealColdFrozenReplayProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('742 ALL12 REAL COLD FIELDS: INDEPENDENT ALL6/WEAK/GAUGE EQUILIBRIUM VERIFIED; NO PHYSICAL LIMIT')
