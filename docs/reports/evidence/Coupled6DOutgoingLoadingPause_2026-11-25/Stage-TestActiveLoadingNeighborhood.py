"""ДЛЯ ТЕСТОВ. Проверяет исходящую координату в заранее заданной окрестности исходного полного прогноза, сохраняя все силовые критерии."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold814ActiveLoadingCoordinate_20261125';root=base/'Cold815ActiveLoadingNeighborhood_20261125'
closure=json.loads((parent/'Native/RealColdReplay.txt.ProcessClosure.json').read_bytes());assert closure['Passed'] and closure['ChildHasExited']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ActiveNeighborhoodCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
method,=re.findall(r'(?ms)^Friend Function TestSolveActiveLoadingCoordinate\(.*?^End Function',code)
anchor='        ByVal maxIterations As Long, ByRef result() As Double, ByRef resultLambda As Double) As Boolean';assert method.count(anchor)==1
updated=method.replace(anchor,'        ByVal maxIterations As Long, ByVal referenceRadius As Double, ByRef result() As Double, ByRef resultLambda As Double) As Boolean',1)
anchor='    Dim projection() As Double, radius As Double: radius = 3# * SpatialFieldDistance(mAccepted, seed)';assert updated.count(anchor)==1
updated=updated.replace(anchor,'    Dim projection() As Double, radius As Double: radius = referenceRadius',1)
code=code.replace(method,updated,1);values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='            mReverseOptionsTest(4) * 0.0001, 15, half, halfLambda)';assert code.count(anchor)==1
code=code.replace(anchor,'            mReverseOptionsTest(4) * 0.0001, 15, 1.5 * mReverseSolverTest.SpatialFieldDistance(accepted, fullSeed), half, halfLambda)',1)
anchor='            mReverseOptionsTest(4) * 0.0001, 30 - CLng(halfEffort(0)), full, fullLambda)';assert code.count(anchor)==1
code=code.replace(anchor,'            mReverseOptionsTest(4) * 0.0001, 30 - CLng(halfEffort(0)), _\n            3# * mReverseSolverTest.SpatialFieldDistance(accepted, fullSeed), full, fullLambda)',1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig');imports=['CCoupledNewtonSolver','modTestCoupledElasticState']
values.update({'Compile-TestExternalPoint.ps1':compiler.replace("-ModuleNames 'modTestCoupledElasticState'",
 '-ModuleNames @('+','.join(repr(n) for n in imports)+')',1).replace('\n','\r\n').encode('utf-8-sig'),
 'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():
 if n.endswith(('.cls','.bas')):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
 (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='OutgoingCoordinateInOriginalPredictorNeighborhood',actual_import_this_pass=imports,actual_component_count=256,
 working_methods_materials_and_source43_fields_exact814=True,
 reference_radius_fixed_before_solving_from_original_full_predictor=True,
 half_radius1_5_full_radius3_times_original_predictor=True,
 target_coordinate_half_of_original_predictor_change=True,
 no_neighborhood_expansion_after_seeing_a_root=True,
 full_and_half_physical_force_gauge_tolerances_unchanged=True,primary_pair30_unchanged=True,
 no_material_acceptance_or_capacity_search=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('815 GEOMETRIC NEIGHBORHOOD FIXED FROM ORIGINAL FULL PREDICTOR; FORCE/GAUGE/PHYSICS/30 UNCHANGED')
