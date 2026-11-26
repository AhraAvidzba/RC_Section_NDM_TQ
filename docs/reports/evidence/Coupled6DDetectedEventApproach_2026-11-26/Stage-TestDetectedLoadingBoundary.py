"""ДЛЯ ТЕСТОВ. Проверяет существующий общий поиск материального события на фактическом полном прогнозе949 без ручного выбора точки."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold949WorkingGuidePair_20261126';root=base/'Cold952DetectedLoadingBoundary_20261126'
assert json.loads((base/'Cold951Point18EventRoots_20261126/Point18EventRootsProof.json').read_bytes())['independently_verified_material_coordinate_equilibria']==2
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
ancestor=root/'DetectedEventCompilationAncestor';ancestor.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(ancestor/name)
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((parent/'Native/WorkingGuidePairs.bin').read_bytes())
assert r.VectorTest().tolist()==[1.];meta,accepted,full,half,projection=[r.VectorTest() for _ in range(5)]
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
updates={};name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n"
anchor=prefix*14+'\n\nPublic Function TestRunWorkingGuidePair()';assert code.count(anchor)==1
entry=(here/'TestRunDetectedLoadingBoundary.bas.txt').read_text('utf-8-sig').rstrip()
updates[name]=code.replace(anchor,prefix*15+'\n\n'+entry+'\n\n\nPublic Function TestRunWorkingGuidePair()',1).replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
updates['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'modTestCoupledElasticState' -EntryPointNames 'TestRunDetectedLoadingBoundary'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig')
old='modTestCoupledElasticState.TestRunWorkingGuidePair';assert runner.count(old)==1
updates[name]=runner.replace(old,'modTestCoupledElasticState.TestRunDetectedLoadingBoundary',1).replace('\n','\r\n').encode('utf-8-sig')
context=[meta[2],meta[0],source['options'][3],source['options'][4],source['options'][2]]
updates.update({'Input/DetectedEvent.bin':vector(context)+vector(full)+vector(projection),
    'SourceWorkingGuidePairs.bin':(parent/'Native/WorkingGuidePairs.bin').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes(),
    'TestRunDetectedLoadingBoundary.bas.txt':(here/'TestRunDetectedLoadingBoundary.bas.txt').read_bytes()})
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():
    if name.endswith('.bas'):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='GenericDetectedEventFromActualFull949Prediction',actual_import_this_pass=['modTestCoupledElasticState'],
    actual_component_count=257,all_numerical_classes_exact949=True,no_hardcoded_point_or_event_type=True,
    existing_working_TryLoadingYieldBoundary_used=True,all129_candidates_eligible_by_existing_rules=True,
    original_force_tolerances_and30_event_steps=True,earlier_prediction_cost_not_included=True,
    no_material_history_acceptance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'GENERIC EXISTING LOADING EVENT API;ONE TEST ENTRY;SAME NUMERICAL CLASSES')
