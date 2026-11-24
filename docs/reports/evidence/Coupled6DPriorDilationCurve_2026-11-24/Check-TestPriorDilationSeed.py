"""ДЛЯ ТЕСТОВ. Проверяет прогноз по прошлой ветви и исходные уравнения локального полного шага."""
from pathlib import Path
from runpy import run_path
import csv,json,re
import numpy as np


def TestCheck(root):
    """Первые11 состояний должны совпасть; далее изменение корня учитывается явно и проверяется по равновесию."""
    here=Path(__file__).resolve().parent
    source=(here/'Check-TestPostYieldLocalProbe.py').read_text('utf-8')
    source=source.replace("for key in ('calibration','options','offset','direction','fields','fractions','lambdas'):assert np.array_equal(a[key],b[key]),key",
        "for key in ('calibration','options','offset','direction'):assert np.array_equal(a[key],b[key]),key\n    for key in ('fields','fractions','lambdas'):assert np.array_equal(a[key][:11],b[key][:11]),key")
    source=source.replace('accepted18_history_and_parameters_exact=True','accepted11_history_exact=True,all_parameters_exact=True')
    source=source.replace("path=root/'PostYieldLocalProbeProof.json'","path=root/'PriorDilationSeedProof.json'")
    namespace={'__file__':str(here/'Check-TestPostYieldLocalProbe.py'),'__name__':'prior_dilation_seed_test'}
    exec(compile(source,str(here/'Check-TestPostYieldLocalProbe.py'),'exec'),namespace)
    proof=namespace['TestCheck'](root)
    fields=run_path(str(here/'Check-TestPostCrackFields.py'))['TestAnalyse'](root/'Matrix','PostCrackPrescribed.Accepted')
    assert fields['states']==18
    a=(root.parent/'Cold709PostYieldLocalProbe_20261124/src/Materials/CCoupledConcreteResponse.cls').read_text('utf-8-sig')
    b=(root/'src/Materials/CCoupledConcreteResponse.cls').read_text('utf-8-sig')
    begin=b.index('    Dim priorTrace As Double, priorQ As Double, priorHigh As Double, priorMiddle As Double, priorLow As Double\n')
    end=b.index('    Dim limit As Double: limit = Tan(mDilationMax)',begin)
    assert b[:begin]+b[end:]==a,'Изменение должно ограничиваться начальным прогнозом дилатансии.'
    reference=proof['reference']
    rows=list(csv.reader((root/'Matrix/Native/LateralFixedIntegration.csv').read_text('cp1251').splitlines(),delimiter=';'))[1:]
    samples=[]
    for row in rows:
        if row[0]!='SAMPLE' or int(row[1]) not in (-1,0) or row[4]!='1':continue
        stress=np.asarray(row[6:12],float)
        if max(abs(stress[:2]))<=1e-10 and abs(float(row[3])-reference['expected_lateral'])<=1e-11 and abs(stress[2]-reference['expected_sigma_zz_MPa'])<=.5/21168:
            samples.append(dict(mode=int(row[1]),index=int(row[2]),lateral=float(row[3]),stress_MPa=stress.tolist()))
    assert samples,'Нужна фактическая нативная проба, согласованная с независимой формулой.'
    # Исходный порог бокового напряжения бискекции 1e-8 не гарантирует дополнительный допуск 0,5Н для продольной силы в почти плоском локальном отклике.
    root_reference_misses=[dict(mode=mode,root=row) for mode in (-1,0) for row in proof['modes'][mode]['roots']
        if row['lateral_error']>1e-11 or row['sigma_zz_error_MPa']>.5/21168]
    proof.update(only_initial_dilation_center_changed=True,all_saved18_equilibria_verified=fields,
        original_full_step_equations_and_acceptance_unchanged=True,analytically_matched_actual_samples=samples,
        bisection_roots_not_meeting_additional_reference_tolerance=root_reference_misses,
        physical_capacity_certified=False,global_postyield_step_not_solved=True)
    (root/'PriorDilationSeedProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('710 VERIFIED SAME EQUATIONS; OPERATIONAL ROOT',proof['operational_full_step_root_found'],'SCALAR ROOT',proof['scalar_full_step_root_found'])
    return proof


if __name__=='__main__':TestCheck(Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold710PriorDilationSeed_20261124'))
