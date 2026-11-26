"""ДЛЯ ТЕСТОВ. Готовит возврат к исходным61 уравнениям из проверенного материального равновесия940."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold936RecordedYieldCoordinate_20261126';source=base/'Cold940RecordedControlBracket_20261126'
root=base/'Cold941OriginalControlReturn_20261126'
proof=json.loads((source/'RecordedYieldCoordinateProof.json').read_bytes())['cases'][0]
assert proof['independently_equilibrated'] and proof['case']==1
chosen=run_path(str(here/'Read-TestRecordedYieldCoordinate.py'))['TestRead'](source/'Native/RecordedYieldCoordinate.bin')[0]
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
ancestor=root/'OriginalControlReturnCompilationAncestor';ancestor.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(ancestor/name)
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((source/'Input/RecordedYieldSeed.bin').read_bytes());r.position=0
meta,current,accepted,residual,projection=[r.VectorTest() for _ in range(5)];assert r.position==len(r.raw)
meta[3]=chosen['info'][6];current=chosen['field'];residual=chosen['values'][:60].copy();residual[56:]/=meta[8]
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
values={};name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n"
anchor=prefix*11+'\n\nPublic Function TestSolveRecordedYieldCoordinate()';assert code.count(anchor)==1
entry=(here/'TestSolveOriginalControlReturn.bas.txt').read_text('utf-8-sig').rstrip()
values[name]=code.replace(anchor,prefix*12+'\n\n'+entry+'\n\n\nPublic Function TestSolveRecordedYieldCoordinate()',1).replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'modTestCoupledElasticState' -EntryPointNames 'TestSolveOriginalControlReturn'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig')
old='modTestCoupledElasticState.TestSolveRecordedYieldCoordinate';assert runner.count(old)==1
values[name]=runner.replace(old,'modTestCoupledElasticState.TestSolveOriginalControlReturn',1).replace('\n','\r\n').encode('utf-8-sig')
values.update({'Input/RecordedYieldSeed.bin':b''.join(vector(v) for v in (meta,current,accepted,residual,projection)),
    'SourceMaterialRoot.bin':(source/'Native/RecordedYieldCoordinate.bin').read_bytes(),
    'VerifiedStartingMaterialRoot.json':json.dumps(proof,sort_keys=True,indent=2).encode('utf-8'),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes(),
    'TestSolveOriginalControlReturn.bas.txt':(here/'TestSolveOriginalControlReturn.bas.txt').read_bytes()})
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in values.items():
    if name.endswith('.bas'):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='OriginalFull61ControlReturnFromMaterialRoot',actual_import_this_pass=['modTestCoupledElasticState'],
    actual_component_count=257,all_numerical_classes_exact936=True,starting_root940_case=1,
    original_projection_target_and_six_force_path_unchanged=True,source_history25_unchanged=True,
    original30_steps_and_force_tolerances=True,material_history_not_accepted=True,
    physical_capacity_certified=False,no_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'ORIGINAL61 EQUATIONS;EXACT MATERIAL ROOT SEED;ONE TEST ENTRY ONLY')
