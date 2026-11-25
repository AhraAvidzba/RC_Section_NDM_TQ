"""ДЛЯ ТЕСТОВ. Различает разгрузочную ветвь, отдельные H-равновесия и ещё не согласованное продолжение через материальный переход."""
from pathlib import Path
from runpy import run_path
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold840EnvelopeReentry_20261125';parent=base/'Cold839ContinuedSpectralWalk_20261125';original=base/'Cold838ElasticPrincipalArc_20261125'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
assert len(source['fields'])==113
for name in ('src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledArcContinuation.cls','src/Solver/CCoupledDeformationContinuation.cls',
    'src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledReinforcedResponse.cls','src/Materials/CCoupledConcreteReturn.cls',
    'src/Solver/CCoupledEnvelopeEquations.cls','src/Solver/CCoupledMaterialConstraint.cls','Input/Reverse1.bin'):
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
name='src/Solver/CCoupledNewtonSolver.cls';old=(original/name).read_text('utf-8-sig');renamed=(parent/name).read_text('utf-8-sig')
renamed=renamed.replace('TestLocateElasticPrincipalBoundary','TryElasticPrincipalBoundary')
marker="' -------------------- ДЛЯ ТЕСТОВ --------------------\n' На неизменной истории выбирает ближайший разрешимый ноль главного напряжения упругого CDP"
assert renamed.count(marker)==1;renamed=renamed.replace(marker,marker.split('\n',1)[1],1)
assert renamed==old,'839 must only rename the test-only entry, without changing working math'
validate=run_path(str(here/'Validate-TestHardPrepared.py'));mechanics,fields=validate['TestValidate'](root,'EnvelopeEnergy.Accepted')
np.testing.assert_allclose(fields,source['fields'],rtol=1e-14,atol=1e-300)
rows=validate['TestRows'](root/'Native/EnvelopeEnergy.Accepted.csv');stored=np.zeros((113,3,43,2))
for row in rows:
    if row[0]=='S':stored[int(row[1])-1,int(row[2])+1,int(row[3])-1]=np.array(row[10:12],float)
coords=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)[4]
gt,gc,length=source['calibration'][6:9];density=np.array([gt/length,gc/length])
h=np.einsum('nspc,p,c->n',stored[:,:,:25],coords[:25,2],density)/(3*sum(coords[:25,2]))
report=validate['TestRows'](root/'Native/EnvelopeReentry.csv');native_h=np.array([float(row[2]) for row in report if row[0]=='H'])
assert len(native_h)==113;np.testing.assert_allclose(h,native_h,rtol=1e-13,atol=1e-18)
last=int(np.flatnonzero(np.max(abs(np.diff(stored,axis=0)),axis=(1,2,3))>0)[-1]+2)
assert last==70 and np.array_equal(stored[last-1:],np.broadcast_to(stored[last-1],stored[last-1:].shape))
assert np.all(np.diff(source['lambdas'][last-1:])<0)
start=[row for row in report if row[0]=='START'];assert len(start)==1
start=np.array(start[0][1:],float);assert start[0]==last and abs(start[3]-h[last-1])<=1e-18
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=Reader((root/'Native/EnvelopeReentry.bin').read_bytes())
assert r.raw[:16]==b'NDMHREENTRY00000'  # Фактически записанные16 байтов первого диагностического формата, без нормализации заголовка.
raw=[];cases={}
for case in range(1,9):
    meta,effort,field,actual=[r.VectorTest() for _ in range(4)]
    assert meta[0]==case and meta[1]==1 and meta[2]==(1 if case<=4 else .0001)
    assert meta[8]==meta[9] and effort[0]<=30 and meta[7]>=1
    assert meta[10] or meta[11]
    raw.append((meta,effort,field,actual));cases[case]=dict(field=field,lambda_value=float(meta[6]),accuracy=float(meta[2]),actual=actual)
assert r.position==len(r.raw)
independent,M=run_path(str(here/'TestUnacceptedMechanics.py'))['TestUnacceptedRoots'](root,source,cases)
results=[]
for case,(meta,effort,field,actual) in enumerate(raw,1):
    mechanics_case=independent[case];energy=mechanics_case['independent_point_energy_average_MPa']
    assert abs(energy-meta[5])<=1e-16 and energy>meta[3]
    assert abs(meta[3]-h[last-1])<=1e-18
    row_scale=source['calibration'][0]/source['calibration'][3]
    assert abs(meta[4]-energy)*row_scale<=mechanics_case['bound_MPa']+1e-12
    expected=np.array([meta[15],meta[14],meta[16]])
    np.testing.assert_allclose(expected,mechanics_case['steel_utilization_minus_center_plus'],rtol=1e-12,atol=1e-15)
    results.append(dict(case=case,lambda_value=float(meta[6]),accuracy_scale=float(meta[2]),baseline_H_MPa=float(meta[3]),
        target_H_MPa=float(meta[4]),actual_H_MPa=float(meta[5]),probes=int(meta[7]),primary_iterations=int(effort[0]),
        mechanics=mechanics_case,forward=bool(meta[10]),native_half_verified=bool(meta[11])))
reference=source['fields'][last-1];factor=source['lambdas'][last-1]
def TestLength(delta):
    """Та же физическая тензорная метрика трёх станций, независимо собранная из операторов."""
    return float(np.sqrt(max(0.,delta@M@delta)))
pairs=[]
for full,half in ((5,6),(6,7),(7,8)):
    full_field=cases[full]['field'];half_field=cases[half]['field'];lf=cases[full]['lambda_value'];lh=cases[half]['lambda_value']
    field_error=TestLength(full_field-(2*half_field-reference))/TestLength(full_field-reference)
    lambda_error=abs(2*lh-factor-lf);lambda_bound=max(.005*abs(lf-factor),8*.0001*source['options'][4]*(1+abs(factor)))
    verified=field_error<=.005 and lambda_error<=lambda_bound
    assert not verified
    pairs.append(dict(full_case=full,half_case=half,field_error=field_error,lambda_error=float(lambda_error),
        lambda_bound=float(lambda_bound),verified=False))
proof=dict(native=native,all113_source_equilibria_verified=mechanics,
    material_energy_variables_not_damage_factors=True,last_material_growth_state=last,
    all129_material_energy_coordinates_identical_after70=True,last43states_have_frozen_material_energies=True,
    later_loads_strictly_decrease=True,full113_H_independently_rebuilt=True,
    all_working_math_and_materials_unchanged=True,test_only_entry_name_corrected=True,
    reentry_roots=results,all8_separate_equilibria_independently_verified=True,strict_full_half_pairs=pairs,
    no_reentry_pair_accepted=True,no_unknown_candidate_material_history_accepted=True,
    requires_local_material_transition_before_H_branch_admission=True,
    physical_capacity_certified=False,whole_goal_completed=False)
(root/'EnvelopeReentryProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('113 SOURCE ROOTS VERIFIED;MATERIAL ENERGY FROZEN AFTER70;8 H ROOTS VERIFIED;ALL3 STRICT PAIRS UNRESOLVED;NO CANDIDATE ACCEPTED')
