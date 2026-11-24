"""ДЛЯ ТЕСТОВ. Подтверждает точные65 полей718 по повтору720 и независимому пространственному равновесию."""
from pathlib import Path
from runpy import run_path
import json
import numpy as np


def TestCheck(root):
    """Повтор без поиска не выдаётся за новый холодный предел; тензоры остаются нативным материалом."""
    here=Path(__file__).resolve().parent
    native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'PostCrackReplay')
    assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
    assert 'sourceStates=65; mode=0;' in text
    frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](root/'Input/PostCrackSourceRawHistory.bin')
    frame=frames[-1]
    raw_input=(root/'Input/PostCrackReplay.bin').read_bytes();assert raw_input[:16]==b'NDMPOSTREPLAY001'
    Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];reader=Reader(raw_input)
    size,=reader.ReadTest('<i');assert size==60
    metadata={key:reader.VectorTest() for key in ('calibration','options','offset','direction')}
    count,=reader.ReadTest('<i');assert count==len(frame['fields'])==65
    for i in range(count):
        fraction,factor=reader.ReadTest('<dd');field=reader.VectorTest()
        assert fraction==frame['fractions'][i] and factor==frame['lambdas'][i] and np.array_equal(field,frame['fields'][i])
    assert reader.position==len(raw_input)
    phase='PostCrackRawReplay.Accepted'
    fields=run_path(str(here/'Check-TestPostCrackFields.py'))['TestAnalyse'](root,phase)
    mechanics,actual=run_path(str(here/'Validate-TestHardPrepared.py'))['TestValidate'](root,phase)
    assert fields['states']==65 and len(actual)==65
    np.testing.assert_allclose(np.asarray(actual),frame['fields'],rtol=1e-14,atol=1e-300)
    for a,b in zip(mechanics['states'],frame['lambdas']):assert abs(a['lambda']-b)<=1e-14*(1+abs(b))
    analytical=run_path(str(here/'Check-TestShoulderAccuracy.py'))['TestAnalyticalForces'](root,phase,fields)
    misses=[v for v in analytical if v['error_N']>.5]
    source=root.parent/'Cold718ControlledLoadBalance_20261124'
    assert (root/'Input/PostCrackSourceRawHistory.bin').read_bytes()==(source/'Native/Stock.RawHistory.bin').read_bytes()
    for name in ('src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledDeformationContinuation.cls','src/Materials/CCoupledConcreteResponse.cls'):
        assert (root/name).read_bytes()==(source/name).read_bytes(),name
    proof=dict(native=native,source_raw_history=raw,all65_input_fields_and_loads_bits_preserved=True,
        all_exported65_fields_and_loads_match_source=True,independent_fields=fields,
        analytical_forces=analytical,analytical_force_misses=misses,
        all_homogeneous_analytical_forces_within_halfNewton=not misses,
        source_material_and_solver_classes718_exact=True,replay_does_not_search_capacity=True,
        physical_capacity_certified=False,independent_material_integration=False,entire_goal_completed=False)
    (root/'PostCrackRawReplayProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('720 EXACT65 REPLAY AND INDEPENDENT EQUILIBRIUM VERIFIED; MAX ANALYTICAL ERROR',
        max(v['error_N'] for v in analytical),'MISSES',len(misses),'LAST',fields['point_fields'][-1],'NO CAPACITY SEARCH')
    return proof


if __name__=='__main__':TestCheck(Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold720PostCrackRawReplay_20261124'))
