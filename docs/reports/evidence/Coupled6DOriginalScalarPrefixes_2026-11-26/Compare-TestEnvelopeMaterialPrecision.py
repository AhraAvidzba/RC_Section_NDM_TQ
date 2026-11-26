"""ДЛЯ ТЕСТОВ. Проверяет, изменило ли ужесточение локальных численных допусков одну и ту же H-пробу."""
from argparse import ArgumentParser
import json
from pathlib import Path
from runpy import run_path
import numpy as np


def TestCompare(base, refined, output=None):
    """Сопоставляет фактические поля и напряжения; исходную принятую историю проверяет отдельно от уточнённых проб."""
    here=Path(__file__).resolve().parent
    check=run_path(str(here/'Check-TestEnvelopeDifferenceBatch.py'))['TestCheck']
    original, fine=check(base),check(refined)
    assert (base/'Input/Reverse1.bin').read_bytes()==(refined/'Input/Reverse1.bin').read_bytes()
    assert original['target_H_MPa']==fine['target_H_MPa']
    rows=run_path(str(here/'Validate-TestHardPrepared.py'))['TestRows']
    comparisons={}
    for index in (1,2,4):
        phase='EnvelopeDifference'+str(index)+'.Sequence'
        before=rows(base/'Native'/(phase+'.csv'));after=rows(refined/'Native'/(phase+'.csv'))
        assert (base/'Native'/(phase+'.Operators.csv')).read_bytes()==(refined/'Native'/(phase+'.Operators.csv')).read_bytes()
        headers=lambda data:{r[0]:r[1:] for r in data if r[0] in ('DIM','LAW','OPTIONS','OFFSET','DIRECTION')}
        assert headers(before)==headers(after)
        first_fields=np.asarray([r[4:] for r in before if r[0]=='F'],float)
        second_fields=np.asarray([r[4:] for r in after if r[0]=='F'],float)
        assert np.array_equal(first_fields[:219],second_fields[:219])
        prefix=lambda data:np.asarray([r[4:] for r in data if r[0]=='S' and int(r[1])<=219],float)
        assert np.array_equal(prefix(before),prefix(after))
        final=lambda data:np.asarray([r[4:] for r in data if r[0]=='S' and int(r[1])==220],float)
        old,new=final(before),final(after)
        comparisons[index]=dict(exact_same_219_fields_and_stress_history=True,
            maximum_final_field_absolute_difference=float(np.max(abs(first_fields[-1]-second_fields[-1]))),
            maximum_final_concrete_stress_difference_MPa=float(np.max(abs(old[:,:6]-new[:,:6]))),
            maximum_final_steel_stress_difference_MPa=float(np.max(abs(old[:,8]-new[:,8]))),
            lambda_absolute_difference=abs(original['cases'][index]['lambda_value']-fine['cases'][index]['lambda_value']),
            original_best_residual_MPa=original['cases'][index]['best_augmented_residual_MPa'],
            refined_best_residual_MPa=fine['cases'][index]['best_augmented_residual_MPa'],
            original_strict_success=original['cases'][index]['strict_solver_success'],
            refined_strict_success=fine['cases'][index]['strict_solver_success'])
    proof=dict(original=original,refined=fine,comparisons=comparisons,
        local_precision_changed_only_after_original_history=True,
        test_cache_includes_exact_precision_factor=True,physical_parameters_unchanged=True,
        independent_material_reintegration=False,physical_capacity_certified=False)
    if output is not None:output.write_bytes((json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode())
    return proof


if __name__=='__main__':
    p=ArgumentParser(description=__doc__)
    for name in ('base','refined','out'):p.add_argument('--'+name,type=Path,required=True)
    a=p.parse_args();proof=TestCompare(a.base.resolve(),a.refined.resolve(),a.out.resolve())
    print(json.dumps(proof['comparisons'],indent=2))
