"""ДЛЯ ТЕСТОВ. Измеряет различие двух чередующихся матриц и геометрическое усиление ошибки напряжений; физический предел не устанавливает."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold907Cold32TangentConeIterations_20261126'
proof=json.loads((root/'Cold32TangentConeIterationsProof.json').read_bytes());assert proof['all12_available']
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];path=root/'Native/Cold32TangentConeIterations.bin'
r=R(path.read_bytes());meta=r.VectorTest();accepted=r.VectorTest();incoming=r.VectorTest();projection=r.VectorTest();records=[]
for _ in range(12):
    info=r.VectorTest();assert info[3]
    probe=r.VectorTest();movement=r.VectorTest();scales=r.VectorTest();residual=r.VectorTest()
    matrix=np.array([r.VectorTest() for _ in range(61)]).T;tangent=r.VectorTest()
    for _ in range(3):r.VectorTest();r.VectorTest()
    records.append(dict(info=info,scales=scales,matrix=matrix,tangent=tangent))
assert r.position==len(r.raw)
a,b=records[:2];np.testing.assert_array_equal(a['scales'],b['scales'])
difference=a['matrix'][:60,:60]-b['matrix'][:60,:60]
singular=np.linalg.svd(difference,compute_uv=False);row_norm=np.linalg.norm(difference,axis=1)
rank=int(np.sum(singular>singular[0]*1e-10)) if singular[0]>0 else 0
interpolated=[]
for theta in (0,.25,.5,.75,1):
    matrix=(1-theta)*a['matrix'][:60,:60]+theta*b['matrix'][:60,:60]
    rhs=-a['matrix'][:60,60]/a['scales'][60]
    spectrum=np.linalg.svd(matrix,compute_uv=False);item=dict(theta=theta,condition=float(spectrum[0]/spectrum[-1]))
    try:
        vector=np.linalg.solve(matrix,rhs)*a['scales'][:60]
        item.update(linear_system_solved=True,linear_direction=vector.tolist(),direction_not_equilibrium=True)
    except np.linalg.LinAlgError:item['linear_system_solved']=False
    interpolated.append(item)
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
gain=np.sum(abs(w),axis=(0,1))+np.sum(abs(v),axis=(0,1))/.001
bound=1.214547470504934e-10
result=dict(source_binary_sha256=sha256(path.read_bytes()).hexdigest().upper(),physical_matrix_difference_singular_values=singular.tolist(),
 numerical_difference_rank_at_relative1e_minus10=rank,largest_difference_rows=(np.argsort(row_norm)[::-1][:8]+1).tolist(),
 difference_row_norms=row_norm.tolist(),interpolated_linear_systems=interpolated,
 triangle_inequality_gain_from_uniform_component_stress_error=gain[:56].tolist(),row53_gain=float(gain[52]),
 reference_arc_equilibrium_bound_MPa=bound,sufficient_uniform_stress_component_error_for_row53_MPa=bound/float(gain[52]),
 scalar_yield_residual_is_not_a_stress_component_error_bound=True,
 no_cause_of_NumFail_proved_by_operator_gain_alone=True,no_material_or_solver_change=True,
 no_new_NDS_or_material_acceptance=True,physical_capacity_certified=False,whole_goal_completed=False)
(root/'Cold32CycleLinearAlgebraProof.json').write_text(json.dumps(result,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('DIFFERENCE RANK',rank,'TOP SINGULAR',singular[:6].tolist(),'TOP ROWS',result['largest_difference_rows'])
print('ROW53 STRESS ERROR GAIN',gain[52],'SUFFICIENT COMPONENT ERROR',result['sufficient_uniform_stress_component_error_for_row53_MPa'])
