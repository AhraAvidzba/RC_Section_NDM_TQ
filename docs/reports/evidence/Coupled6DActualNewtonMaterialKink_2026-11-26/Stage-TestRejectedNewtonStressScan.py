"""ДЛЯ ТЕСТОВ. Читает локальные напряжения в реальной неудачной точке Newton и вдоль её поправки; новую ветвь не принимает."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
capture=base/'Cold911NewtonRejectedCapture_20261126';parent=base/'Cold900Cold32OutgoingJacobian_20261126'
root=base/'Cold913RejectedNewtonStressScan_20261126'
proof=json.loads((capture/'NewtonRejectedCaptureProof.json').read_bytes());assert proof['baseline8_cases_and_all_root_exports_bit_identical']
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((capture/'Native/NewtonRejected.00001.bin').read_bytes());r.position=0
meta,current,accepted,residual,rhs,delta=[r.VectorTest() for _ in range(6)]
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Input/Reverse1.bin')
np.testing.assert_array_equal(accepted,source['fields'][-1]);assert len(source['fields'])==32
fractions=[None,0.,1.,1/128,1/256,1/512,1/1024,1/4096,1/16384,1/65536,-1/128,-1/1024,-1/65536]
predictions=[accepted if f is None else current+f*delta[:60] for f in fractions]
factor=source['lambdas'][-1]
cases=[dict(case=i+1,name='accepted_zero' if f is None else 'actual_Newton_fraction_'+str(f),
    newton_fraction=f,lambda_increment=0. if f is None else float(meta[2]+f*delta[60]-factor)) for i,f in enumerate(fractions)]
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
prep=json.loads((root/'Preparation.json').read_bytes());name='tools/build_workbook/Run-Audit02ModeTests.ps1'
runner=(base/'Cold857SpatialDeclineRecovery_20261125'/name).read_text('utf-8-sig')
anchor='        else { $result = [string]$excel.Run($macroName) }';assert runner.count(anchor)==1
runner=runner.replace(anchor,anchor+'''
        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected rejected-Newton stress-scan setup macro.'}
        $null=$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestReadUnresolvedArcScan")
        $lines.Add('READONLY_PRIVATE_SCAN_COMPLETED: no new equilibrium search')
        Save-Progress
''',1)
values={name:runner.replace('\n','\r\n').encode('utf-8-sig'),
 'Input/ArcScanFields.bin':vector([len(predictions)])+b''.join(vector(f) for f in predictions),
 'Input/ArcScanCases.json':(json.dumps(cases,sort_keys=True,indent=2)+'\n').encode(),
 'Input/ActualRejectedNewton.bin':(capture/'Native/NewtonRejected.00001.bin').read_bytes(),
 'SourceNewtonRejectedCaptureProof.json':(capture/'NewtonRejectedCaptureProof.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ActualRejectedNewtonFrozenLocalStressScan',actual_import_this_pass=json.loads((parent/'docs/regression/CodeUpdate.json').read_bytes())['Modules'],
 actual_imports_this_replay=[],actual_component_count=256,immutable_compiled_book_reused=True,compiled_source_root=parent.name,
 actual_history32_and_rejected_iterate_exact=True,endpoint_fractions=fractions,
 only_existing_readonly_scan_entry=True,no_new_NDS_search_or_material_acceptance=True,
 physical_capacity_certified=False,no_working_or_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
print(root,'EXACT COMPILED900;13 FROZEN ENDPOINTS;NO NEW NDS SEARCH')
