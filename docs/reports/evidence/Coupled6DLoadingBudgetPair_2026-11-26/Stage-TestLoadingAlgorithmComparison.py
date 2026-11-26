"""ДЛЯ ТЕСТОВ. Сравнивает существующие Newton/dogleg на положительной материальной ветви с отдельным бюджетом30, без принятия состояния."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold978LoadingMaterialDirection_20261126';root=base/'Cold979LoadingAlgorithmComparison_20261126'
proof=json.loads((parent/'MaterialDirectionProof.json').read_bytes())
assert proof['physical_capacity_certified'] is False
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
name='tools/build_workbook/Run-Audit02ModeTests.ps1';code=(root/name).read_text('utf-8-sig')
old='modTestCoupledElasticState.TestRunMaterialDirectionAudit';assert code.count(old)==1
updates={name:code.replace(old,'modTestCoupledElasticState.TestSolveRecordedYieldCoordinate',1).replace('\n','\r\n').encode('utf-8-sig'),
    Path(__file__).name:Path(__file__).read_bytes()}
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='PositiveMaterialCoordinateStandaloneNewtonDogleg30',actual_import_this_pass=[],actual_component_count=257,
    all257_actual_components_and_saved_book_bitexact978=True,no_new_numerical_code_or_compilation=True,
    same_positive_full_half_targets_and_actual_failed968_seed=True,original30_per_independent_case=True,
    not_a_shared_budget_full_half_pair=True,no_material_history_acceptance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'EXACT978 BOOK;EXISTING NEWTON/DOGLEG30;NOT JOINT PAIR OR CAPACITY')
