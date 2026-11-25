"""ДЛЯ ТЕСТОВ. Восстанавливает фактический префикс113 из исходных Double и продолжает неизменную спектральную дугу."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold838ElasticPrincipalArc_20261125';root=base/'Cold839ContinuedSpectralWalk_20261125'
proof=json.loads((parent/'AutomaticLoadingWalkProof.json').read_bytes())
assert proof['new_accepted_equilibria']==27 and proof['total_calls']==95
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Input/Reverse1.bin')
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
fields=list(source['fields']);factors=list(source['lambdas'])
for item in proof['attempts']:
    r=Reader((parent/'Native'/f"LoadingWalk.{item['attempt']:03}.bin").read_bytes())
    assert r.raw[:16]==b'NDMOPERARC000004';packet=[r.VectorTest() for _ in range(15)]
    if item['solved']:
        field,actual,projection=[r.VectorTest() for _ in range(3)]
        assert packet[0][23]==item['lambda_value'];fields.append(field);factors.append(packet[0][23])
    r.VectorTest();r.VectorTest();r.VectorTest();assert r.position==len(r.raw)
assert len(fields)==113
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ContinuedSpectralCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
assert code.count('Friend Function TryElasticPrincipalBoundary(')==1
assert code.count('TryElasticPrincipalBoundary = FollowYieldEventCoordinate')==1
code=code.replace('Friend Function TryElasticPrincipalBoundary(','Friend Function TestLocateElasticPrincipalBoundary(',1)
code=code.replace('TryElasticPrincipalBoundary = FollowYieldEventCoordinate','TestLocateElasticPrincipalBoundary = FollowYieldEventCoordinate',1)
anchor="' На неизменной истории выбирает ближайший разрешимый ноль главного напряжения упругого CDP, не принимая найденное поле."
assert code.count(anchor)==1
code=code.replace(anchor,"' -------------------- ДЛЯ ТЕСТОВ --------------------\n"+anchor,1)
values[name]=code
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
assert code.count('mReverseSolverTest.TryElasticPrincipalBoundary(')==1
code=code.replace('mReverseSolverTest.TryElasticPrincipalBoundary(','mReverseSolverTest.TestLocateElasticPrincipalBoundary(',1)
values[name]=code
values['Input/Reverse1.bin']=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](source,
    dict(fields=np.array(fields),fractions=np.ones(len(fields)),lambdas=np.array(factors)))
imports=['CCoupledNewtonSolver','modTestCoupledElasticState']
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
assert compiler.count("-ModuleNames 'modTestCoupledElasticState'")==1
values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'",'-ModuleNames @('+','.join(repr(n) for n in imports)+')',1)
values.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():
    if isinstance(data,str):
        if name.endswith(('.bas','.cls')):assert max(map(len,data.splitlines()))<1024
        data=data.replace('\n','\r\n').encode('utf-8-sig')
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ContinueActual113StateSpectralBranch',actual_import_this_pass=sorted(imports),actual_component_count=256,
    exact113_prefix_reconstructed_from_original_838_packets=True,
    only_test_entry_name_changed_in_solver=True,all_working_solver_math_and_materials_exact838=True,
    maximum_calls96=True,maximum_attempts32=True,maximum_new_points32=True,
    all_three_station_steel_checks_preserved=True,no_prescribed_future_field=True,
    not_cold_capacity_search=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root,'EXACT113;UNCHANGED EQUATIONS;TEST ENTRY NAMED EXPLICITLY')
