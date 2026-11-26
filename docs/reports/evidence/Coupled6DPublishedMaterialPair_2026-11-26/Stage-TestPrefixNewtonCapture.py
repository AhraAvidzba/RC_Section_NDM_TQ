"""ДЛЯ ТЕСТОВ. Включает существующую запись отклонённого Newton на точной книге967 без изменения численных методов."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold967LongerSecantPrefix_20261126';root=base/'Cold968PrefixNewtonCapture_20261126'
proof=json.loads((parent/'ControlledSecantPrefixProof.json').read_bytes());assert proof['returned_verified_candidate'] and proof['root_step_budget']==10
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
name='tools/build_workbook/Run-Audit02ModeTests.ps1';code=(root/name).read_text('utf-8-sig')
anchor='            $lines.Add([string]$excel.Run("\'RC_Section_NDM.xlsm\'!modTestCoupledElasticState.TestRunControlledSecantPrefix"))'
assert code.count(anchor)==1
insert='            $null = $excel.Run("\'RC_Section_NDM.xlsm\'!modTestColdNewtonCapture.TestSetCaptureRoot", $MacroArgument1)\n'
data=code.replace(anchor,insert+anchor,1).replace('\n','\r\n').encode('utf-8-sig')
(root/name).write_bytes(data);prep=json.loads((root/'Preparation.json').read_bytes());prep['inputs'][name]=sha256(data).hexdigest().upper()
(root/Path(__file__).name).write_bytes(Path(__file__).read_bytes());prep['inputs'][Path(__file__).name]=sha256(Path(__file__).read_bytes()).hexdigest().upper()
prep['directed_scope'].update(case='ExistingRejectedNewtonCaptureOnSameSecantPrefix967',actual_import_this_pass=[],
    all257_actual_components_and_saved_book_bitexact967=True,no_new_numerical_code_or_compilation=True,
    no_extra_material_call_for_capture=True,captures_only_already_computed_rejected_systems=True,
    original_shared30_and_root10_and_force_half_steel_criteria_preserved=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'EXACT257-BOOK967;EXISTING NEWTON CAPTURE ONLY;NO EXTRA MATERIAL EVALUATION')
