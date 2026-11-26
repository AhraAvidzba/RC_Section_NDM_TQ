"""ДЛЯ ТЕСТОВ. Проверяет половинный подход к автоматически найденному событию952 на исходных допусках и общем бюджете30."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold952DetectedLoadingBoundary_20261126';root=base/'Cold953DetectedEventApproach_20261126'
proof=json.loads((parent/'DetectedLoadingBoundaryProof.json').read_bytes());assert proof['independently_equilibrated'] and proof['effort'][0]==2
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
ancestor=root/'EventApproachCompilationAncestor';ancestor.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(ancestor/name)
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
face=np.array(proof['field']);accepted=source['fields'][-1]
context=[source['lambdas'][-1],proof['lambda_value'],source['options'][3],source['options'][4],source['options'][2],proof['effort'][0]]
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
updates={};name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n"
anchor=prefix*15+'\n\nPublic Function TestRunDetectedLoadingBoundary()';assert code.count(anchor)==1
entry=(here/'TestRunDetectedEventApproach.bas.txt').read_text('utf-8-sig').rstrip()
updates[name]=code.replace(anchor,prefix*16+'\n\n'+entry+'\n\n\nPublic Function TestRunDetectedLoadingBoundary()',1).replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
updates['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'modTestCoupledElasticState' -EntryPointNames 'TestRunDetectedEventApproach'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig')
old='modTestCoupledElasticState.TestRunDetectedLoadingBoundary';assert runner.count(old)==1
updates[name]=runner.replace(old,'modTestCoupledElasticState.TestRunDetectedEventApproach',1).replace('\n','\r\n').encode('utf-8-sig')
updates.update({'Input/ResolvedEventApproach.bin':vector(context)+vector(face)+vector(accepted),
    'SourceDetectedEventProof.json':(parent/'DetectedLoadingBoundaryProof.json').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes(),
    'TestRunDetectedEventApproach.bas.txt':(here/'TestRunDetectedEventApproach.bas.txt').read_bytes()})
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():
    if name.endswith('.bas'):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='StandardToleranceHalfApproachToGenericEvent',actual_import_this_pass=['modTestCoupledElasticState'],
    actual_component_count=257,all_numerical_classes_exact952=True,standard_force_tolerances_and_half_limit_preserved=True,
    event2_steps_plus_half_within_total30=True,existing_spatial_direction_control_and_CurveLambdaBound_used=True,
    no_material_history_acceptance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'EVENT2 STEPS + HALF, TOTAL30;STANDARD TOLERANCES;.5% GUARD;NO ACCEPTANCE')
