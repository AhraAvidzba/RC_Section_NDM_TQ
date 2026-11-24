"""ДЛЯ ТЕСТОВ. Проверяет численный исход одного прямого шага отдельно от успеха диагностического макроса."""
from pathlib import Path
from argparse import ArgumentParser
from runpy import run_path
from hashlib import sha256
import json, re
import numpy as np


def TestCheck(root):
    """Сохраняет отрицательную аналитическую сверку, даже когда слабое равновесие и нативный макрос успешны."""
    here=Path(__file__).resolve().parent
    native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'PostCrackReplay')
    assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
    count,=re.findall(r'sourceStates=(\d+); mode=1;',text);count=int(count)
    outcome,=re.findall(r'DIRECT_PROBE: solved=(True|False);',text)
    frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](root/'Input/PostCrackSourceRawHistory.bin')
    source=frames[-1];assert count<=len(source['fields'])
    proof=dict(native=native,numerical_step_solved=outcome=='True',source_states=count,
        physical_capacity_certified=False,new_cold_capacity_search=False,independent_material_integration=False)
    if outcome=='True':
        phase='PostCrackRawReplay.Sequence'
        fields=run_path(str(here/'Check-TestPostCrackFields.py'))['TestAnalyse'](root,phase)
        mechanics,actual=run_path(str(here/'Validate-TestHardPrepared.py'))['TestValidate'](root,phase)
        assert fields['states']==count+1
        np.testing.assert_allclose(np.asarray(actual)[:count],source['fields'][:count],rtol=1e-14,atol=1e-300)
        analytical=run_path(str(here/'Check-TestShoulderAccuracy.py'))['TestAnalyticalForces'](root,phase,fields)
        proof.update(independent_fields=fields,last_field=fields['point_fields'][-1],
            last_analytical=analytical[-1],last_analytical_within_original_halfNewton=analytical[-1]['error_N']<=.5,
            analytical_forces=analytical,analytical_force_misses=[p for p in analytical if p['error_N']>.5],
            exported_source_prefix_matches=True)
    proof['script_sha256']=sha256(Path(__file__).read_bytes()).hexdigest().upper()
    (root/'DirectProbeProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('DIRECT STEP',outcome,'SOURCE',count,'LAST ANALYTICAL',proof.get('last_analytical'),'NO CAPACITY CERTIFICATE')
    return proof


if __name__=='__main__':
    p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,required=True)
    TestCheck(p.parse_args().root.resolve())
