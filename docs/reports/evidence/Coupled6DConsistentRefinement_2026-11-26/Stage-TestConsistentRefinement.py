"""ДЛЯ ТЕСТОВ. Разрешает уже существующее уточнение касательной после неверного направления consistentControl; материалы и допуски сохраняет."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import argparse,json

parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('--baseline',action='store_true');args=parser.parse_args()
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold895FreshSpatialArc_20261126';source=base/'Cold898FreshArcColdStress_20261126'
root=base/('Cold902ConsistentRefinementBaseline_20261126' if args.baseline else 'Cold901ConsistentControlRefinement_20261126')
assert json.loads((base/'Cold900Cold32OutgoingJacobian_20261126/Cold32OutgoingJacobianProof.json').read_bytes())['outgoing_working_predictor']['native_TrySpatialLoadTangent_probe_is_exact_1_over1024']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ConsistentRefinementCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
old='                If Not mConsistentControl And Not mRefinedControlTangent And SupportsHybridRecovery And mIterations + 1 < maxIterations Then'
new='                If Not mRefinedControlTangent And SupportsHybridRecovery And mIterations + 1 < maxIterations Then'
assert code.count(old)==1
if not args.baseline:code=code.replace(old,new,1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
entry=(here/'TestCold32RefinementSteps.bas.txt').read_text('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='Private Const HARD_PREPARED_HALF_TEST As Double = 0.001\n';assert code.count(anchor)==1
values[name]=code.replace(anchor,anchor+"' FOR TESTS: baseline entry point only; existing methods retained.\n\n\n"+entry,1).replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames @('CCoupledNewtonSolver','modTestCoupledElasticState') -EntryPointNames 'TestReadCold32RefinementSteps'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(base/'Cold857SpatialDeclineRecovery_20261125'/name).read_text('utf-8-sig')
anchor='        else { $result = [string]$excel.Run($macroName) }';assert runner.count(anchor)==1
runner=runner.replace(anchor,anchor+'''
        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected consistent-refinement setup macro.'}
        $lines.Add([string]$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestReadCold32RefinementSteps"))
        Save-Progress
''',1)
values[name]=runner.replace('\n','\r\n').encode('utf-8-sig')
values.update({'Input/Reverse1.bin':(source/'Input/Reverse1.bin').read_bytes(),
 'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ConsistentControlRefinementBaseline' if args.baseline else 'ConsistentControlRefinementCandidate',
 actual_import_this_pass=['CCoupledNewtonSolver','modTestCoupledElasticState'],actual_component_count=256,
 only_working_change_remove_Not_mConsistentControl_from_existing_mismatch_upgrade=not args.baseline,
 baseline_all_working_exact895=args.baseline,original_force_material_and_iteration_budgets=True,
 source_states=32,independent_probe_ratios=[1,.5,.25],accepted_history_not_advanced=True,
 physical_capacity_certified=False,no_working_or_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'BASELINE' if args.baseline else 'ONE EXISTING REFINEMENT GATE CORRECTED','SAME STRICT TOLERANCES AND30')
