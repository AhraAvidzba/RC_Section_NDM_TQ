"""ДЛЯ ТЕСТОВ. Проверяет рабочий решатель на материальной координате точки11 вместо неудачной линейной проекции, от той же истории25."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import re,json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold935Point11Matrices_20261126';root=base/'Cold936RecordedYieldCoordinate_20261126'
proof=json.loads((parent/'SurfaceDirectionSignsProof.json').read_bytes())
assert any(x['level']>0 and x['directional_trial_yield_derivative_MPa']<0 for x in proof['directions'])
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((parent/'Input/ActualRejectedNewton.bin').read_bytes());r.position=0
meta,current,accepted,residual,rhs,delta=[r.VectorTest() for _ in range(6)]
A=np.array([r.VectorTest() for _ in range(61)]);end=r.VectorTest();n=int(r.VectorTest()[0])
probes=[[r.VectorTest() for _ in range(3)] for _ in range(n)];assert r.position==len(r.raw)
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Input/Reverse1.bin')
np.testing.assert_array_equal(accepted,source['fields'][-1])
remaining=json.loads((Path.cwd()/'docs/reports/evidence/Coupled6DSurfaceDescent_2026-11-26/RemainingSurfaceOriginProof.json').read_bytes())
station=remaining['surface_crossings_at_minimum_fraction'][0];assert station['point']==11 and station['side']==-1
target_after=station['trial_yield_MPa'];vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
event_meta=[11.,-1.,0.,meta[2],meta[5],source['options'][3],source['options'][4],source['options'][2],meta[4],target_after]
fixture=b''.join(vector(v) for v in (event_meta,current,accepted,residual,A[-1,:60]/meta[4]))
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'RecordedYieldCoordinateCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n"
anchor=prefix*10+'\n\nPublic Function TestReadSurfaceMatrices() As String';assert code.count(anchor)==1
entry=(here/'TestSolveRecordedYieldCoordinate.bas.txt').read_text('utf-8-sig')
first=re.search(r'(?ms)^Public Function TestSolveRecordedYieldCoordinate\(.*?^End Function\b',entry).group()
second=re.search(r'(?ms)^Public Sub TestWriteRecordedYieldStressTest\(.*?^End Sub\b',entry).group()
replacement=prefix*11+'\n\n'+first+'\n\n'+prefix+'\n\n'+second+'\n\n\nPublic Function TestReadSurfaceMatrices() As String'
values[name]=code.replace(anchor,replacement,1).replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'modTestCoupledElasticState' -EntryPointNames 'TestSolveRecordedYieldCoordinate','TestWriteRecordedYieldStressTest'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig')
assert runner.count('modTestCoupledElasticState.TestReadSurfaceMatrices')==1
values[name]=runner.replace('modTestCoupledElasticState.TestReadSurfaceMatrices','modTestCoupledElasticState.TestSolveRecordedYieldCoordinate',1).replace('\n','\r\n').encode('utf-8-sig')
values.update({'Input/RecordedYieldSeed.bin':fixture,
 'SourcePoint11DirectionSignsProof.json':(parent/'SurfaceDirectionSignsProof.json').read_bytes(),
 'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():
    if name.endswith('.bas'):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='Point11MaterialCoordinateReplacesLinearProjection',actual_import_this_pass=['modTestCoupledElasticState'],
 actual_component_count=257,all_numerical_classes_exact935=True,accepted_history25_unchanged=True,
 original_material_and_six_component_load_path=True,scalar_control_intentionally_changed_to_trial_yield=True,
 original_linear_control_solution_not_claimed=True,targets=[0.,target_after],solver_algorithms=['complete Newton line search','hybrid dogleg'],
 every_independent_case_limited_to_original30_steps=True,source_residual_bitexact_guard_before_solve=True,
 material_history_not_accepted=True,physical_capacity_certified=False,no_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'TWO MATERIAL TARGETS,TWO OWN SOLVER MODES;SAME MATERIAL/SIX FORCE PATH;NO HISTORY ACCEPTANCE')
