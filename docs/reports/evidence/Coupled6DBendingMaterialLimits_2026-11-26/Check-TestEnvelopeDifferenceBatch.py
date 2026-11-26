"""ДЛЯ ТЕСТОВ. Проверяет сравнение разностных стратегий на общей H-цели и независимо пересобирает принятые кандидаты."""
from argparse import ArgumentParser
import json
from pathlib import Path
import re
from runpy import run_path
import struct
import numpy as np


def TestCheck(root, output=None):
    """Разделяет успех строгого решателя, пригодность по исходным критериям и ещё не подтверждённый физический предел."""
    here = Path(__file__).resolve().parent
    validate = run_path(str(here/'Validate-TestHardPrepared.py'))
    rows = validate['TestRows'](root/'Native/IntegratedEvent.csv')
    head = {r[0]: r[1:] for r in rows if r[0] in ('START', 'END')}
    count=int(head['START'][0])
    case_count=int(head['END'][0])
    assert count >= 147 and case_count in (4,5)
    assert head['END'][2] == head['END'][3]
    raw = (root/'Input/EnvelopeDifferenceOptions.bin').read_bytes()
    size, = struct.unpack_from('<i', raw)
    assert size == 3*case_count and len(raw) == 4+8*size
    options = np.frombuffer(raw, dtype='<f8', offset=4).reshape(case_count, 3)
    assert np.all(np.isfinite(options)) and np.all((options[:,0]>=1e-8)&(options[:,0]<=.001))
    assert set(options[:,1]).issubset({0.,1.}) and set(options[:,2]).issubset({1.,.1,.01})
    solves = {int(r[1]): r for r in rows if r[0] == 'SOLVE'}
    native = {int(r[1]): r for r in rows if r[0] == 'NATIVE'}
    best = {int(r[1]): r for r in rows if r[0] == 'BEST'}
    residual = {int(r[1]): np.asarray(r[2:], float) for r in rows if r[0] == 'RESIDUAL'}
    assert set(solves) == set(native) == set(range(1,case_count+1))
    original = run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
    assert len(original['fields'])==count
    baseline_proof,baseline_values=validate['TestValidate'](root,'EnvelopeEnergy.Accepted')
    assert len(baseline_values)==count
    np.testing.assert_allclose(baseline_values,original['fields'],rtol=1e-14,atol=1e-300)
    full_target = float(head['START'][3])
    levels={int(r[1]):r for r in rows if r[0]=='LEVEL'}
    if case_count==5:
        assert set(levels)==set(solves)
        fractions=[.125,.25,.5,.75,1.]
        baseline=float(head['START'][2])
        for index,fraction in enumerate(fractions,1):
            assert abs(float(levels[index][2])-(baseline+fraction*(full_target-baseline)))<1e-17
    fields = {}
    cases = {}
    reference_prefix = None
    for index in range(1,case_count+1):
        target=float(levels[index][2]) if levels else full_target
        solve, check = solves[index], native[index]
        assert float(solve[2]) == options[index-1,0]
        assert solve[3] == ('-1' if options[index-1,1] else '0') and float(solve[4]) == options[index-1,2]
        solved = solve[5] == '-1'
        tolerance = float(solve[8])
        if index in residual:
            assert residual[index].shape == (61,)
            assert abs(max(abs(residual[index])) - float(best[index][3])) < 1e-14
            assert abs(float(solve[7]) - float(best[index][3])) < 1e-12
            if solved: assert max(abs(residual[index])) <= tolerance + 1e-14
        phase = 'EnvelopeDifference'+str(index)+'.Sequence'
        path = root/'Native'/(phase+'.csv')
        independent = None
        if path.exists():
            assert check[2:5] == ['-1','-1','-1']
            equilibrium, values = validate['TestValidate'](root, phase)
            assert len(values) == count+1
            np.testing.assert_allclose(values[:count], original['fields'], rtol=1e-14, atol=1e-300)
            if reference_prefix is None: reference_prefix = values[:count]
            else: assert np.array_equal(reference_prefix, values[:count])
            assert abs(equilibrium['states'][-1]['lambda'] - float(check[6])) < 1e-13
            records = validate['TestRows'](path)
            headers = {r[0]:r[1:] for r in records if r[0] in ('DIM','LAW')}
            dofs, _, points, concrete = map(int, headers['DIM'][:4])
            law = np.asarray(headers['LAW'],float)
            _,_,_,_,coords,_,b,g = run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native'/(phase+'.Operators.csv'),dofs,points)
            energy = np.zeros((3,concrete,2)); seen = np.zeros((3,concrete),int)
            for r in records:
                if r[0]=='S' and int(r[1])==count+1 and int(r[3])<=concrete:
                    side,point=int(r[2])+1,int(r[3])-1
                    energy[side,point]=np.asarray(r[10:12],float);seen[side,point]+=1
            assert np.all(seen==1)
            h=float(np.einsum('sp,p->',(law[6]*energy[...,0]+law[7]*energy[...,1])/law[8],coords[:concrete,2])/(3*sum(coords[:concrete,2])))
            assert abs(h-float(check[7])) < 1e-16
            original_tolerance=tolerance/options[index-1,2]
            assert abs(h-target)*law[0]/law[3] <= original_tolerance+1e-12
            fields[index]=values[-1]
            independent=dict(equilibrium=equilibrium,last_H_MPa=h,original_H_verified=True)
        matrix_analysis=None
        matrix_path=root/'Native'/('HomotopyMatrix'+str(index)+'.csv')
        if matrix_path.exists():
            matrix_rows=validate['TestRows'](matrix_path)
            columns={int(r[1]):np.asarray(r[2:],float) for r in matrix_rows if r[0]=='COLUMN'}
            assert set(columns)==set(range(1,62))
            matrix=np.column_stack([columns[j] for j in range(1,62)])
            assert matrix.shape==(61,61) and np.all(np.isfinite(matrix))
            singular=np.linalg.svd(matrix,compute_uv=False)
            step=np.linalg.lstsq(matrix,-residual[index],rcond=None)[0]
            matrix_analysis=dict(singular_values=singular.tolist(),condition=float(singular[0]/singular[-1]),
                linearized_best_residual_MPa=float(max(abs(matrix@step+residual[index]))),
                normalized_Newton_step_norm=float(np.linalg.norm(step)),linear_solve_is_not_equilibrium=True)
        cases[index]=dict(relative_difference_step=float(solve[2]),complete_newton=solve[3]=='-1',
            internal_tolerance_ratio=float(solve[4]),strict_solver_success=solved,stop_code=int(solve[6]),
            best_augmented_residual_MPa=float(solve[7]),strict_tolerance_MPa=tolerance,
            iterations=int(solve[9]),function_calls=int(solve[10]),jacobian_calls=int(solve[11]),seconds=float(solve[12]),
            unchanged_original_force_candidate=check[2:5]==['-1','-1','-1'],
            spatial_advance_verified_by_native=check[5]=='-1',lambda_value=float(check[6]),
            independent_candidate=independent,target_H_MPa=target,matrix_analysis=matrix_analysis)
    assert len(fields)==int(head['END'][1])
    result=json.loads((root/'Native/ExternalResult.json').read_bytes())
    owner=json.loads((root/'Native/ExternalOwnership.json').read_bytes())
    assert result['source_unchanged'] and result['owned_excel_final_state']['state']=='Exited'
    assert result['owned_excel_final_state']['filetime']==owner['filetime']
    summary=dict(cases=cases,initial_states=count,exported_original_force_candidates=sorted(fields),
        same_accepted_history_for_all_variants=True,accepted_revision_unchanged_during_all_probes=True,
        target_H_MPa=full_target,own_excel_closed=True,independent_material_reintegration=False,
        accepted_history_independent_equilibrium=baseline_proof,
        diagnostic_candidates_are_not_certified_capacity=True,physical_capacity_certified=False)
    if output is not None: output.write_bytes((json.dumps(summary,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode())
    return summary


if __name__=='__main__':
    p=ArgumentParser(description=__doc__)
    p.add_argument('--root',type=Path,required=True);p.add_argument('--out',type=Path,required=True)
    a=p.parse_args();result=TestCheck(a.root.resolve(),a.out.resolve())
    print(json.dumps({k:{a:b for a,b in v.items() if a not in ('independent_candidate','matrix_analysis')} for k,v in result['cases'].items()},indent=2))
