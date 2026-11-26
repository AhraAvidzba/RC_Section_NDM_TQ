"""ДЛЯ ТЕСТОВ. Читает активную сторону точного отказа Newton с сохранением исходных скалярной строки и столбца нагрузки."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold911NewtonRejectedCapture_20261126';root=base/'Cold914ProspectiveBranchMatrix_20261126'
assert json.loads((base/'Cold913RejectedNewtonStressScan_20261126/ActualNewtonKinkDiagnosisProof.json').read_bytes())['actual_failing_step_not_accepted_state_predictor']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ProspectiveBranchCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n"
anchor=prefix*8+'\n\nPublic Sub TestNewtonRejectedCapture('
assert code.count(anchor)==1
entry=(here/'TestProspectiveBranchMatrix.bas.txt').read_text('utf-8-sig').rstrip()
code=code.replace(anchor,prefix*9+'\n\n'+entry+'\n\n\nPublic Sub TestNewtonRejectedCapture(',1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'modTestCoupledElasticState' -EntryPointNames 'TestProspectiveBranchMatrix'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(base/'Cold857SpatialDeclineRecovery_20261125'/name).read_text('utf-8-sig')
anchor='        else { $result = [string]$excel.Run($macroName) }';assert runner.count(anchor)==1
runner=runner.replace(anchor,anchor+'''
        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected prospective matrix setup macro.'}
        $lines.Add([string]$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestProspectiveBranchMatrix"))
        Save-Progress
''',1);values[name]=runner.replace('\n','\r\n').encode('utf-8-sig')
raw=(parent/'Native/NewtonRejected.00001.bin').read_bytes()
import struct
position=0
for _ in range(6+61):
    n,=struct.unpack_from('<i',raw,position);position+=4+8*n
values.update({'Input/FrozenNewton.bin':raw[:position],
 'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():
    if name.endswith('.bas'):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ActualFailedNewtonProspectiveActiveBranch',actual_import_this_pass=['modTestCoupledElasticState'],
 actual_component_count=256,all_working_sources_exact911=True,exact_original_control_row_and_load_column=True,
 no_material_acceptance=True,no_new_full_NDS_or_capacity_search=True,
 test_read_context_scalar_row_ignored_and_replaced_by_captured_exact_row=True,
 physical_capacity_certified=False,no_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'ACTUAL CURRENT;PROSPECTIVE MATRICES;ORIGINAL CONTROL/LOAD COLUMN EXACT;ONE TEST ENTRY ONLY')
