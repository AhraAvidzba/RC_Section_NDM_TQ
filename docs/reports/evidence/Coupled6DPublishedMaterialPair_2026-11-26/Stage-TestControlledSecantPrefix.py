"""ДЛЯ ТЕСТОВ. Подготавливает префиксы фактической неудачной меры по принятой секущей без будущего равновесия."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold960ResolvedFieldProgress_20261126';root=base/'Cold965ControlledSecantPrefix_20261126'
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
r=R((base/'Cold935Point11Matrices_20261126/Input/ActualRejectedNewton.bin').read_bytes());r.position=0
meta,current,accepted,residual,rhs,delta=[r.VectorTest() for _ in range(6)]
matrix=np.array([r.VectorTest() for _ in range(61)])
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Input/Reverse1.bin')
np.testing.assert_array_equal(accepted,source['fields'][-1]);p=matrix[-1,:60]/meta[4]
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
ancestor=root/'SecantPrefixCompilationAncestor';ancestor.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(ancestor/name)
updates={};name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
block=(here/'TestTryControlledSecantPrefix.cls.txt').read_text('utf-8-sig').rstrip()
updates[name]=(code+'\n\n'+block+'\n').replace('\n','\r\n').encode('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n"
anchor=prefix*22+'\n\nPublic Function TestRunFieldProgressSteps()';assert code.count(anchor)==1
entry=(here/'TestRunControlledSecantPrefix.bas.txt').read_text('utf-8-sig').rstrip()
updates[name]=code.replace(anchor,prefix*23+'\n\n'+entry+'\n\n\nPublic Function TestRunFieldProgressSteps()',1).replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
updates['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'CCoupledNewtonSolver','modTestCoupledElasticState' -EntryPointNames 'TestRunControlledSecantPrefix'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig')
old='modTestCoupledElasticState.TestRunFieldProgressSteps';assert runner.count(old)==1
updates[name]=runner.replace(old,'modTestCoupledElasticState.TestRunControlledSecantPrefix',1).replace('\n','\r\n').encode('utf-8-sig')
vec=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
context=[meta[5],source['options'][3],source['options'][4],source['options'][2],source['lambdas'][-1],1.]
updates['Input/ControlledPrefix.bin']=vec(context)+vec(p)
updates.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes(),
    'TestTryControlledSecantPrefix.cls.txt':(here/'TestTryControlledSecantPrefix.cls.txt').read_bytes(),
    'TestRunControlledSecantPrefix.bas.txt':(here/'TestRunControlledSecantPrefix.bas.txt').read_bytes()})
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():
    if name.endswith(('.bas','.cls')):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='OriginalFailedScalarPrefixesFromAcceptedSecant',actual_import_this_pass=['CCoupledNewtonSolver','modTestCoupledElasticState'],
    actual_component_count=257,all255_other_components_exact960=True,working_newton_methods_exact960=True,
    test_only_adapter_to_existing_full_half_return=True,no_supplied_future_equilibrium=True,
    original25_history_and_force_path_and_materials_exact960=True,shared30_steps_for_all_attempts_and_full_half=True,
    original_force_tolerances_and_half_005_and_steel_and_resolved_field_guards_preserved=True,
    diagnostic_operation_budget_not_remaining_budget_of_full_cold=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'ORIGINAL SCALAR;ACCEPTED SECANT ONLY;REAL SHARED30;EXISTING FULL/HALF VALIDATION;TEST ADAPTER ONLY')
