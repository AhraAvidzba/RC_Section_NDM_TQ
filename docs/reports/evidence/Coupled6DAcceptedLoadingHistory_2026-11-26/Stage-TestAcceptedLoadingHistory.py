"""ДЛЯ ТЕСТОВ. Добавляет штатное принятие подтверждённой нагружающей пары и два последующих шага от обновляемой истории."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold980LoadingBudgetPair_20261126';root=base/'Cold984AcceptedLoadingHistory_20261126'
proof=json.loads((parent/'MaterialDirectionProof.json').read_bytes())
assert proof['equilibrated_cases_with_resolved_H_growth']==2 and proof['ordinary_elastic_unloading_cases']==0
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
ancestor=root/'AcceptedHistoryCompilationAncestor';ancestor.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(ancestor/name)
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
begin=code.index('Public Function TestRunPublishedMaterialPair()');end=code.index('\nEnd Function',begin)+len('\nEnd Function')
pair=code[begin:end].replace('TestRunPublishedMaterialPair','TestRunHistoryMaterialPair')
old='If mReverseFieldsTest.Count <> 25 Then';assert pair.count(old)==1
pair=pair.replace(old,'If mReverseFieldsTest.Count < 25 Then',1)
pair=pair.replace('Нужна точная принятая история25.','Нужна принятая история не короче исходных 25 точек.',1)
pair=pair.replace('от одной истории25','от текущей принятой истории',1)
entry=(here/'TestRunAcceptedLoadingHistory.bas.txt').read_text('utf-8-sig').rstrip()
anchor='Public Function TestRunMaterialDirectionAudit() As String';assert code.count(anchor)==1
updates={name:code.replace(anchor,entry+'\n\n'+pair+'\n\n'+anchor,1).replace('\n','\r\n').encode('utf-8-sig'),
    'TestRunAcceptedLoadingHistory.bas.txt':(entry+'\n').replace('\n','\r\n').encode('utf-8-sig'),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
updates['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig')
old='modTestCoupledElasticState.TestRunMaterialDirectionAudit';assert runner.count(old)==1
updates[name]=runner.replace(old,'modTestCoupledElasticState.TestRunAcceptedLoadingHistory',1).replace('\n','\r\n').encode('utf-8-sig')
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():
    if name.endswith('.bas'):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='NativeAcceptedLoadingHistoryAndTwoUpdatedSteps',actual_import_this_pass=['modTestCoupledElasticState'],
    actual_component_count=257,all256_numerical_components_exact980=True,only_new_test_entries_added=True,
    original_loading_pair980_reused_before_acceptance=True,native_AcceptStep_and_same_field_post_commit_check=True,
    at_most_three_steps_each_original_shared30=True,subsequent_seed_extrapolated_from_two_accepted_fields=True,
    no_future_equilibrium_field_input=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'NATIVE ACCEPT/POST-COMMIT/UPDATED HISTORY;ALL256 NUMERICAL COMPONENTS EXACT980')
