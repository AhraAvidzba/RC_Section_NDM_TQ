"""ДЛЯ ТЕСТОВ. Проверяет исходную кривую и четыре точки до существующего предела стали0,02."""
from pathlib import Path
from runpy import run_path
import json, csv, math
import numpy as np


def TestCheck(root):
    """Достижение стали в предписанной точке не объявляется завершённым автоматическим поиском."""
    here=Path(__file__).resolve().parent
    native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'CurveMatrix')
    assert native==dict(passed=230,failed=0,macros=1,timeout600=False,own_process_closed=True)
    parent=root.parent/'Cold722PriorBranchClosure_20261124'
    for path in ('src/Materials/CCoupledConcreteResponse.cls','src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledDeformationContinuation.cls'):
        assert (root/path).read_bytes()==(parent/path).read_bytes(),path
    rows=list(csv.reader((root/'Matrix/Native/PostCrackTension.Curve.csv').read_text('cp1251').splitlines(),delimiter=';'))[1:]
    result=[]
    for row in rows:
        e,n,reported,sigma,util=map(float,row)
        sc=30000*e if e<=.0001 else 3*math.exp(-3125*(e-.0001))
        expected=21168*sc+432*min(200000*e,400)
        assert abs(n-expected)<=.5 and abs(reported-expected)<1e-7 and abs(sc-sigma)<1e-12 and abs(util-e/.02)<1e-12
        result.append(dict(epsilon=e,force_N=n,expected_N=expected,error_N=abs(n-expected),steel_utilization=util))
    assert len(result)==25 and [p['epsilon'] for p in result[-4:]]==[.006,.01,.015,.02]
    fields=run_path(str(here/'Check-TestPostCrackFields.py'))['TestAnalyse'](root/'Matrix','PostCrackPrescribed.Accepted')
    assert fields['states']==25 and fields['independent_equilibrium']['states'][-1]['steel_utilization']==1
    read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary']
    new=read(root/'Matrix/Native/PostCrackPrescribed.Accepted25.bin');old=read(parent/'Matrix/Native/PostCrackPrescribed.Accepted21.bin')
    for key in ('calibration','options','offset','direction'):assert np.array_equal(new[key],old[key]),key
    for key in ('fields','fractions','lambdas'):assert np.array_equal(new[key][:21],old[key]),key
    progress=list(csv.reader((root/'Matrix/Native/DilationProgress.csv').read_text('cp1251').splitlines(),delimiter=';'))
    end=progress[-1];assert end[0]=='END' and float(end[1])==.02
    proof=dict(native=native,curve=result,independent_fields=fields,original21_history_bits_exact=True,
        unchanged_material_and_solvers=True,steel_limit_unchanged=.02,steel_strain_limit_reached=True,
        prescribed_solver_elapsed_seconds=float(end[2]),elapsed_excludes_export_and_workbook_start=True,
        physical_capacity_certified=False,automatic_capacity_not_tested=True,entire_goal_completed=False)
    (root/'PriorBranchSteelEndProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('725 PRESCRIBED25 VERIFIED230/0; STEEL LIMIT .02 REACHED; MAX ANALYTICAL ERROR',
        max(p['error_N'] for p in result),'ELAPSED EXCLUDING EXPORT',float(end[2]),'NOT AUTOMATIC CAPACITY')
    return proof


if __name__=='__main__':TestCheck(Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold725SteelEndDirectoryRetry_20261124'))
