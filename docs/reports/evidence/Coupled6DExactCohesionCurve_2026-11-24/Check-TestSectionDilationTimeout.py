"""ДЛЯ ТЕСТОВ. Проверяет тайм-аут707 и законченные строки частичного CSV, не подменяя ими проверку полного поля."""
from pathlib import Path
from runpy import run_path
import json,math
from hashlib import sha256


def TestCheck(root):
    """Явно сохраняет незавершённую строку и отсутствие экспорта поля; исходы частичных сил не дают общий PASS."""
    here=Path(__file__).resolve().parent
    native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'CurveMatrix')
    assert native['timeout600'] and native['own_process_closed'] and native['macros']==0
    path=root/'Matrix/Native/PostCrackTension.Curve.csv';raw=path.read_bytes();lines=raw.decode('cp1251').splitlines(keepends=True)
    valid=[];incomplete=[]
    for i,line in enumerate(lines[1:],2):
        if not line.endswith(('\r','\n')):
            incomplete.append(dict(line=i,text=line));continue
        row=[float(v) for v in line.strip().split(';')];assert len(row)==5 and all(math.isfinite(v) for v in row)
        epsilon,force,reported,stress,util=row
        sigma=30000*epsilon if epsilon<=.0001 else 3*math.exp(-3125*(epsilon-.0001))
        expected=21168*sigma+432*min(200000*epsilon,400)
        assert abs(force-expected)<=.5 and abs(reported-expected)<1e-7 and abs(stress-sigma)<1e-12 and abs(util-epsilon/.02)<1e-12
        valid.append(dict(epsilon=epsilon,recorded_force_N=force,expected_force_N=expected,error_N=abs(force-expected)))
    assert valid and incomplete and valid[-1]['epsilon']>=.0008
    assert not (root/'Matrix/Native/PostCrackPrescribed.Accepted.csv').exists()
    proof=dict(native=native,completed_CSV_rows_independently_checked=valid,incomplete_rows=incomplete,
        last_complete_recorded_strain=valid[-1]['epsilon'],maximum_force_error_N=max(v['error_N'] for v in valid),
        complete_field_equilibrium_not_exported=True,complete_curve_not_verified=True,
        target10seconds_not_met=True,physical_capacity_certified=False,entire_goal_completed=False,
        raw_partial_curve_sha256=sha256(raw).hexdigest().upper())
    (root/'SectionDilationTimeoutProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('707TIMEOUT VERIFIED; COMPLETE CSV ROWS',len(valid),'LAST',valid[-1]['epsilon'],'INCOMPLETE',incomplete,'NO FULL FIELD PASS')
    return proof


if __name__=='__main__':TestCheck(Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold707SectionDilationClosure_20261124'))
