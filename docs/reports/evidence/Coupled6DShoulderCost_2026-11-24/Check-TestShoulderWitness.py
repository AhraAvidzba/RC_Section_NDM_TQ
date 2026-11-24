"""ДЛЯ ТЕСТОВ. Различает успех макроса, подтверждение точности и независимую проверку конечного поля плеча."""
from pathlib import Path
from argparse import ArgumentParser
from runpy import run_path
import csv,json,re


def TestCheck(root):
    """Не выдаёт незавершённую выгрузку или аналитическое расхождение за подтверждённый предел."""
    here=Path(__file__).resolve().parent
    native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'PostCrackReplay')
    assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
    outcome=(root/'Native/QualityOutcome.txt').read_text('cp1251')
    a,=re.findall(r'CRITICAL_QUALITY_PROBE: directSolved=(True|False); qualityConfirmed=(True|False);',outcome)
    assert ('directSolved='+a[0]+'; qualityConfirmed='+a[1]+';') in text
    initial,final,probes=re.findall(r'initialUncertainty=([^;]+); finalUncertainty=([^;]+); additionalProbes=(\d+);',outcome)[0]
    profile=list(csv.reader((root/'Native/ReplayProfile.csv').read_text('cp1251').splitlines(),delimiter=';'))
    retries=[dict(attempt=int(p[1]),elapsed=float(p[2]),accuracy_scale=float(p[3]),uncertainty_before=float(p[4])) for p in profile if p[0]=='QUALITY_RETRY']
    end,=[p for p in profile if p[0]=='DIRECT_END']
    proof=dict(native=native,direct_solved=a[0]=='True',quality_confirmed=a[1]=='True',
        initial_uncertainty=float(initial),final_uncertainty=float(final),additional_probes=int(probes),
        original_quality_target=.001/16,quality_retries=retries,elapsed_before_export_seconds=float(end[2]),
        physical_capacity_certified=False,entire_goal_completed=False,independent_material_integration=False)
    if proof['quality_confirmed']:
        assert proof['final_uncertainty']<=proof['original_quality_target']
        fields=run_path(str(here/'Check-TestPostCrackFields.py'))['TestAnalyse'](root,'PostCrackRawReplay.Sequence')
        assert fields['states']==1
        analytical=run_path(str(here/'Check-TestShoulderAccuracy.py'))['TestAnalyticalForces'](root,'PostCrackRawReplay.Sequence',fields)
        state=fields['independent_equilibrium']['states'][0]
        scale=retries[-1]['accuracy_scale'] if retries else 1.
        proof.update(independent_final_field=fields,analytical=analytical[0],
            analytical_within_original_halfNewton=analytical[0]['error_N']<=.5,
            requested_strict_residual_tolerance_MPa=state['bound_MPa']*scale,
            independent_residual_below_requested_strict_tolerance=(state['residual_MPa']<=state['bound_MPa']*scale))
    (root/'ShoulderWitnessProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('SHOULDER',proof['quality_confirmed'],'U',initial,'->',final,'SECONDS',float(end[2]),'ANALYTIC',proof.get('analytical'),'NOT CAPACITY')
    return proof


if __name__=='__main__':
    p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,required=True)
    TestCheck(p.parse_args().root.resolve())
