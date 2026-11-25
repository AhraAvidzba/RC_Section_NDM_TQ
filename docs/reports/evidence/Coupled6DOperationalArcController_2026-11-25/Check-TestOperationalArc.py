"""ДЛЯ ТЕСТОВ. Независимо проверяет рабочую дуговую операцию785: историю, дугу, событие, половинный подход и фактический34-й корень."""
from pathlib import Path
from runpy import run_path
import json
import numpy as np
from argparse import ArgumentParser

here = Path(__file__).resolve().parent; base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__)
p.add_argument('--root',type=Path,default=base/'Cold785OperationalArcController_20261125')
root=p.parse_args().root
native, _ = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative'](root, 'RealColdReplay')
assert native == dict(passed=1, failed=0, macros=1, timeout600=False, own_process_closed=True)
source = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary'](root / 'Input/Reverse1.bin')
assert len(source['fields']) == 33
assert (root / 'Input/Reverse1.bin').read_bytes() == (base / 'Cold784AutomaticEventRefinement_20261125/Input/Reverse1.bin').read_bytes()
Reader = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']
r = Reader((root / 'Native/OperationalArc.Result.bin').read_bytes()); assert r.raw[:16] == b'NDMOPERARC000001\x00'[:16]
packet = [r.VectorTest() for _ in range(14)]
returned, actual, returned_projection, effort = [r.VectorTest() for _ in range(4)]
assert r.position == len(r.raw)
info, projection, accepted, half_seed, full_seed, half, full, ah, af, event, face, face_half, unused_face_actual, ah_face = packet
assert len(info) == 27 and all(info[:4] == 1) and info[26] == 1
assert info[14] == info[15] == event[6] == 35
assert info[16] <= 30 and info[17] <= 30 and info[18] == 5 <= 6 and info[19] == info[16] + info[17]
assert effort[0] == 0 and np.array_equal(accepted, source['fields'][-1]) and np.array_equal(returned, face)
assert info[6] == source['lambdas'][-1] and info[7] == source['lambdas'][-2] and info[8] == 2
operators = run_path(str(here / 'Check-TestHardHalfSteps.py'))['TestOperators'](root / 'Native/EnvelopeEnergy.Accepted.Operators.csv', 60, 43)
geometry, loads, w, v, coordinates, axis, b, gradient = operators
op = np.array([b + s*.001*gradient for s in (-1,0,1)])
metric = np.einsum('spcd,spce,p,c->de',op,op,coordinates[:,2],np.array([1.,1.,1.,.5,.5,.5]))/(3*sum(coordinates[:,2]))
metric[56:,:] = 0; metric[:,56:] = 0


def TestLength(delta):
    """Восстанавливает тензорную норму из исходных операторов и фактических площадей всех точек."""
    return float(np.sqrt(max(0., delta@metric@delta)))


delta = accepted - source['fields'][-2]; dl = info[6] - info[7]
slopes = [TestLength(source['fields'][i+1]-source['fields'][i])/abs(source['lambdas'][i+1]-source['lambdas'][i]) for i in range(29,32)]
alpha = float(np.median(slopes)); secant = np.sqrt(TestLength(delta)**2 + alpha**2*dl**2)
np.testing.assert_allclose(info[9:13], [alpha,secant,2*secant,alpha**2*dl/secant],rtol=1e-9,atol=1e-23)
np.testing.assert_allclose(projection,metric@delta/secant,rtol=1e-8,atol=1e-17)
np.testing.assert_allclose(half_seed[:56],accepted[:56]+delta[:56],rtol=1e-14,atol=1e-300)
np.testing.assert_allclose(full_seed[:56],2*half[:56]-accepted[:56],rtol=1e-14,atol=1e-300)
scales = np.array([1/loads[0,0],1/loads[1,1],1/loads[2,2],1/loads[3,3],1/loads[0,0],1/loads[0,0]])
force_norms = {}
for label, field, factor, force in (('pair_full',full,info[4],af),('pair_half',half,info[5],ah),('event',returned,info[23],actual),('event_half',face_half,info[24],ah_face)):
    bound = .0001*(source['options'][3]+source['options'][4]*max(abs((loads@(source['offset']+factor*source['direction']))[:56])))
    norm = float(max(abs((force-source['offset']-factor*source['direction'])/scales)))
    gauge = float(max(abs(geometry[56:]@field))*source['calibration'][0])
    assert norm <= bound+1e-14 and gauge <= bound+1e-13
    force_norms[label] = dict(raw_six_residual_MPa=norm,gauge_MPa=gauge,bound_MPa=float(bound))
spacing = TestLength(returned-accepted)
error = TestLength(returned-(2*face_half-accepted))/spacing
np.testing.assert_allclose(error,info[25],rtol=1e-6,atol=1e-14); assert error <= .005
np.testing.assert_allclose(returned_projection, metric@(returned-accepted)/spacing, rtol=1e-8,atol=1e-17)
# Параметры точки15/+1 сняты независимо776 на том же побитном33-точечном префиксе.
assert event[1] == 15 and event[2] == 1
rr = Reader((base/'Cold776ArcPointRoute_20261125/Native/ArcPoint15.Route.bin').read_bytes())
while True:
    identity,strain,sigma = [rr.VectorTest() for _ in range(3)]
    vectors = [rr.VectorTest() for _ in range(9)]
    if identity[0] == 1 and identity[1] == 1:
        committed, parameters = vectors[3], vectors[8]; break
diagonal, off, shear, alpha_c, gamma, *_ = parameters
elastic = np.diag([diagonal]*3+[shear]*3);elastic[:3,:3] += off*(np.ones((3,3))-np.eye(3))


def TestYield(field):
    """Строит пробную поверхность от захваченных до события пластических координат и огибающих."""
    xx,yy,zz,xy,xz,yz = elastic@((b[14]+.001*gradient[14])@field-committed[12:18])
    tensor=np.array([[xx,xy,xz],[xy,yy,yz],[xz,yz,zz]]);trace=xx+yy+zz
    high=np.linalg.eigvalsh(tensor)[-1];q=np.sqrt(1.5*np.sum((tensor-np.eye(3)*trace/3)**2))
    ct,cc=committed[20:22]
    term=(cc/ct*(1-alpha_c)-(1+alpha_c))*high if high>0 else gamma*high
    return float((alpha_c*trace+q+term)/(1-alpha_c)-cc)


np.testing.assert_allclose([TestYield(accepted),TestYield(full)],event[3:5],rtol=1e-6,atol=2e-12)
assert abs(TestYield(face)) <= 2.1e-12 and TestYield(face_half) < 0
mechanics,fields = run_path(str(here/'Validate-TestHardPrepared.py'))['TestValidate'](root,'OperationalArc.Full.Sequence')
assert len(fields) == 34
np.testing.assert_allclose(fields[:33],source['fields'],rtol=1e-14,atol=1e-300)
np.testing.assert_allclose(fields[33],returned,rtol=1e-14,atol=1e-300)
proof=dict(native=native,source33_exact784=True,operational_class_called_without_test_search_driver=True,
    accepted_revision_unchanged=True,nonlinear_calls=int(info[18]),pair_primary_steps=int(info[16]),event_primary_steps=int(info[17]),
    all_primary_steps=int(info[19]),tensor_half_error=float(error),event_lambda=float(info[23]),independent_event_yield_MPa=TestYield(face),
    independent_half_yield_MPa=TestYield(face_half),force_checks=force_norms,all34_independent_weak_and_six_force_equilibria=mechanics,
    controller_cold_start_not_claimed=True,physical_capacity_certified=False,whole_goal_completed=False)
(root/'OperationalArcProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('785 OPERATIONAL CLASS VERIFIED;34 INDEPENDENT EQUILIBRIA;5CALLS/7PRIMARY; NO HISTORY ACCEPTED BY CLASS; NO COLD CAPACITY CLAIM')
