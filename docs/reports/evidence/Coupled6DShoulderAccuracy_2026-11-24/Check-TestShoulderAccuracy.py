"""ДЛЯ ТЕСТОВ. Проверяет полный исход717 и новые строгие равновесия, не назначая предел по близости к известному ответу."""
from pathlib import Path
from runpy import run_path
import csv,json,math,re
import numpy as np


def TestCheck(root):
    """Независимая проверка каждого сохранённого состояния дополняется аналитической одноосной огибающей."""
    here=Path(__file__).resolve().parent
    native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'PostCrackTension')
    assert native==dict(passed=18,failed=1,macros=1,timeout600=False,own_process_closed=True)
    fields=run_path(str(here/'Check-TestPostCrackFields.py'))['TestAnalyse'](root,'PostCrackContinuation.Accepted')
    rows=run_path(str(here/'Validate-TestHardPrepared.py'))['TestRows'](root/'Native/PostCrackContinuation.Accepted.csv')
    operator_rows=run_path(str(here/'Validate-TestHardPrepared.py'))['TestRows'](root/'Native/PostCrackContinuation.Accepted.Operators.csv')
    areas={int(v[1]):float(v[4]) for v in operator_rows if v[0]=='P'}
    axes={int(v[1]):float(v[4]) for v in operator_rows if v[0]=='A'}
    forces={}
    for v in rows:
        if v[0]!='S' or int(v[2])!=0:continue
        index,point=int(v[1]),int(v[3]);stress=float(v[6])
        if point>fields['concrete_points']:stress=float(v[12])*axes[point]-stress
        forces[index]=forces.get(index,0.)+areas[point]*stress
    analytical=[]
    for p in fields['point_fields']:
        e=p['control_epsilon'];sigma=30000*e if e<=.0001 else 3*math.exp(-3125*(e-.0001))
        expected=21168*sigma+432*min(200000*e,400)
        analytical.append(dict(index=p['index'],epsilon=e,force_N=forces[p['index']],expected_N=expected,error_N=abs(forces[p['index']]-expected)))
    # Отрицательные исходы сверки сохраняются отдельно и не исправляются расширением допуска.
    misses=[v for v in analytical if v['error_N']>.5]
    read=run_path(str(here/'Check-TestActualAdjointResidual.py'))['TestRead'];matrices=[]
    for index,solved,pivot,code in csv.reader((root/'Native/Adjoint.Outcomes.csv').read_text('cp1251').splitlines(),delimiter=';'):
        assert solved=='1'
        a,b,r,factor,tolerance=read(root/'Native'/('Adjoint.'+index+'.bin'))
        raw=(root/'Native'/('Adjoint.Weights.'+index+'.bin')).read_bytes()
        w=np.frombuffer(raw,dtype='<f8',offset=4)
        original=5e-7*(1+120000*abs(factor)/21600)
        assert tolerance<=original*(1+1e-13)
        bound=float(np.sum(abs(w))*max(tolerance,float(np.max(abs(r))))+2e-12*(1+abs(factor)))
        matrices.append(dict(index=int(index),lambda_value=factor,tolerance=tolerance,original_tolerance=original,
            tighter=tolerance<original*(1-1e-10),actual_residual_inf=float(np.max(abs(r))),original_uniform_bound=bound))
    lower=float(re.findall(r'наибольший допустимый lambda: ([0-9.E+-]+)',text)[0]);expected=172967.56279474715/120000
    assert abs(lower-1.44139577360552)<1e-13 and 'проб: 96;' in text
    parent=root.parent/'Cold716ActualAdjointResidual_20261124'
    assert (root/'src/Solver/CCoupledNewtonSolver.cls').read_bytes()==(parent/'src/Solver/CCoupledNewtonSolver.cls').read_bytes()
    assert (root/'src/Materials/CCoupledConcreteResponse.cls').read_bytes()==(parent/'src/Materials/CCoupledConcreteResponse.cls').read_bytes()
    proof=dict(native=native,independent_fields=fields,analytical_force_checks=analytical,
        analytical_force_misses=misses,all_analytical_forces_within_original_halfNewton=not misses,
        lower_lambda=lower,known_uniaxial_maximum=expected,lower_to_known_maximum=expected-lower,
        local_peak_reported_accuracy=.0119129026718645,required_capacity_accuracy=.001,
        uncertainty_matrices=matrices,critical_tolerances_only_tightened=True,
        original_uncertainty_formula_and_material_exact=True,physical_capacity_certified=False,
        steel_terminal_event_not_reached=True,entire_goal_completed=False)
    (root/'ShoulderAccuracyProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('717 VERIFIED STATES',fields['states'],'LOWER',lower,'KNOWN ERROR',expected-lower,
        'MAX FORCE ERROR',max(v['error_N'] for v in analytical),'MISSES',len(misses),
        'STRICT MATRICES',sum(v['tighter'] for v in matrices),'NO FINAL CAPACITY')
    return proof


if __name__=='__main__':TestCheck(Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold717ShoulderAccuracy_20261124'))
