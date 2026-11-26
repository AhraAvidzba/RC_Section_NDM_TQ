"""ДЛЯ ТЕСТОВ. Проверяет собственные прогнозные направления и реальные направленные производные без объявления их равновесиями."""
from pathlib import Path
from runpy import run_path
from fractions import Fraction
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold907Cold32TangentConeIterations_20261126';parent=base/'Cold895FreshSpatialArc_20261126'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=2,failed=0,macros=2,timeout600=False,own_process_closed=True)
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
assert audit['components']==256 and audit['imported_this_pass']==['modTestCoupledElasticState']
assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==255
for path in (parent/'src').rglob('*'):
    if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes(),str(path)
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
assert len(source['fields'])==32
validate=run_path(str(here/'Validate-TestHardPrepared.py'));mechanics,fields=validate['TestValidate'](root,'EnvelopeEnergy.Accepted')
np.testing.assert_allclose(fields,source['fields'],rtol=1e-14,atol=1e-300)
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)])
def TestLength(field):
    """Вычисляет ту же тензорную норму напрямую по геометрическим операторам, без нативного метода расстояния."""
    strain=np.einsum('spcd,d->spc',op,field)
    return float(np.sqrt(np.einsum('spc,spc,p,c->',strain,strain,coords[:,2],np.array([1,1,1,.5,.5,.5]))/(3*sum(coords[:,2]))))
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((root/'Native/Cold32TangentConeIterations.bin').read_bytes())
assert r.raw[:16]==b'NDMCONETAN320001'
meta=r.VectorTest();accepted=r.VectorTest();incoming=r.VectorTest();projection=r.VectorTest()
assert len(meta)==8 and meta[0]==3 and meta[1]==4 and meta[6]==61
assert meta[2]==source['lambdas'][-1] and meta[3]==source['lambdas'][-2]
np.testing.assert_array_equal(accepted,source['fields'][-1]);np.testing.assert_array_equal(incoming,accepted-source['fields'][-2])
def TestRoundIntermediate(value,precision):
    """Точно моделирует округление рационального двоичного результата к заданной мантиссе перед записью Double; допуск не вводит."""
    if not value:return 0.
    sign=1 if value>0 else -1;value=abs(value)
    exponent=value.numerator.bit_length()-value.denominator.bit_length()
    if value<Fraction(2)**exponent:exponent-=1
    unit=Fraction(2)**(exponent-precision+1)
    return float(sign*round(value/unit)*unit)
results=[];previous_moves={};zeros=[];arithmetic_rounding=[]
while r.position<len(r.raw):
    info=r.VectorTest();group,denominator,iteration,ready=info[:4]
    assert group in (1,2,3) and denominator==(1024,64,8)[int(group)-1] and iteration in (1,2,3,4)
    if not ready:
        results.append(dict(group=int(group),denominator=denominator,iteration=int(iteration),matrix_available=False));break
    assert len(info)==7
    probe=r.VectorTest();prior=r.VectorTest();scales=r.VectorTest();residual=r.VectorTest()
    matrix=np.array([r.VectorTest() for _ in range(61)]).T;tangent=r.VectorTest()
    if iteration==1:previous_moves[group]=incoming
    np.testing.assert_array_equal(prior,previous_moves[group])
    direct=accepted+prior/denominator
    for component in np.flatnonzero(probe!=direct):
        exact=Fraction.from_float(float(accepted[component]))+Fraction.from_float(float(prior[component]))/Fraction.from_float(float(denominator))
        rounded=TestRoundIntermediate(exact,64)
        assert probe[component]==rounded
        arithmetic_rounding.append(dict(group=int(group),iteration=int(iteration),component=int(component)+1,
            native_equals_exact64bit_mantissa_then_Double=True,direct_Double_result=float(direct[component]),native=float(probe[component])))
    expected=-L@source['direction']*scales[60];expected[56:]*=30000
    np.testing.assert_allclose(matrix[:60,60],expected,rtol=1e-13,atol=1e-13)
    independent=np.linalg.solve(matrix[:60,:60],-matrix[:60,60]/scales[60])*scales[:60]
    rel=float(np.max(abs(independent-tangent))/max(np.max(abs(tangent)),1e-300));assert rel<1e-8
    length=TestLength(tangent);assert abs(length/info[5]-1)<1e-11
    orientation=meta[5]**2*(meta[2]-meta[3])
    for i in range(60):orientation+=projection[i]*tangent[i]*meta[4]
    assert abs((orientation-info[6])/info[6])<1e-11
    scale=float(np.sign(orientation)*meta[4]/length);assert abs(scale/info[4]-1)<1e-11
    movement=info[4]*tangent;previous_moves[group]=movement
    direction=np.r_[movement,info[4]]/scales;calculated=matrix@direction
    item=dict(group=int(group),denominator=denominator,iteration=int(iteration),matrix_available=True,
        lambda_increment=float(info[4]),independent_linear_relative_error=rel,
        linear_predicted_physical_derivative_MPa=float(max(abs(calculated[:56]))),samples=[])
    origin=None
    for beta in (0,1/1024,1/4096):
        header=r.VectorTest();actual=r.VectorTest();assert header.tolist()==[beta,1.]
        if beta==0:origin=actual;zeros.append(actual[:60])
        else:
            derivative=(actual-origin)/beta
            item['samples'].append(dict(beta=beta,actual_physical_derivative_MPa=float(max(abs(derivative[:56]))),
                actual_row53_derivative_MPa=float(derivative[52]),
                derivative_defect_MPa=float(max(abs((derivative-calculated)[:56])))))
    results.append(item)
assert r.position==len(r.raw)
drift=float(np.max(abs(np.asarray(zeros)-zeros[0]))) if zeros else None
if drift is not None:assert drift<1e-11
proof=dict(native=native,actual_components=256,all_working_sources_exact895=True,true_history32_unchanged=True,
 all_available_linear_tangents_independently_solved=True,zero_reading_drift_MPa=drift,readings=results,
 generated_probe_arithmetic_exactly_explained=arithmetic_rounding,raw_history_keys_still_bit_exact=True,
 all12_available=len(results)==12 and all(x['matrix_available'] for x in results),
 no_new_NDS_or_material_acceptance=True,diagnostic_probes_are_not_accepted_history=True,
 physical_capacity_certified=False,whole_goal_completed=False)
(root/'Cold32TangentConeIterationsProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('CONE TANGENT READS',len(results),'ZERO DRIFT',drift,'NO NDS/FAILURE CLAIM')
for item in results:
    print('GROUP',item['group'],'ITERATION',item['iteration'],'AVAILABLE',item['matrix_available'],'LAMBDA STEP',item.get('lambda_increment'))
    if item['matrix_available']:
        print(' LINEAR',item['linear_predicted_physical_derivative_MPa'],'ACTUAL',item['samples'])
