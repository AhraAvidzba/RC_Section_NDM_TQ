"""ДЛЯ ТЕСТОВ. Читает локальные отклики по обе стороны действительной поправки54 из холодного отказа, без новых равновесий."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold925RecordedTailSystem_20261126';source=base/'Cold923RecordedNewtonCold96_20261126'
root=base/'Cold926RecordedTailStressScan_20261126'
proof=json.loads((parent/'RecordedTailSystemProof.json').read_bytes())
assert proof['base_residual_max_error_MPa']==0 and proof['refined_matrix_max_relative_row_error']<2e-12
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((source/'Native/NewtonRejected.00054.bin').read_bytes());r.position=0
meta,current,accepted,residual,rhs,delta=[r.VectorTest() for _ in range(6)]
binary=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Input/Reverse1.bin')
assert len(binary['fields'])==25 and accepted.tobytes()==binary['fields'][-1].tobytes();factor=binary['lambdas'][-1]
fractions=[None,0.,1.,.5,.25,.125,.0625,.03125,.015625,.0078125,1/256,1/1024,1/4096,1/16384,1/65536,-1/128,-1/4096,-1/65536]
predictions=[accepted if f is None else current+f*delta[:60] for f in fractions]
cases=[dict(case=i+1,name='accepted_zero' if f is None else 'actual_tail_Newton_fraction_'+str(f),
    newton_fraction=f,lambda_increment=0. if f is None else float(meta[2]+f*delta[60]-factor)) for i,f in enumerate(fractions)]
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
prep=json.loads((root/'Preparation.json').read_bytes());name='tools/build_workbook/Run-Audit02ModeTests.ps1'
runner=(base/'Cold857SpatialDeclineRecovery_20261125'/name).read_text('utf-8-sig')
anchor='        else { $result = [string]$excel.Run($macroName) }';assert runner.count(anchor)==1
runner=runner.replace(anchor,anchor+'''
        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected recorded tail stress-scan setup macro.'}
        $null=$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestReadUnresolvedArcScan")
        $lines.Add('READONLY_PRIVATE_SCAN_COMPLETED: no new equilibrium search')
        Save-Progress
''',1)
values={name:runner.replace('\n','\r\n').encode('utf-8-sig'),
 'Input/ArcScanFields.bin':vector([len(predictions)])+b''.join(vector(f) for f in predictions),
 'Input/ArcScanCases.json':(json.dumps(cases,sort_keys=True,indent=2)+'\n').encode(),
 'Input/ActualRejectedNewton.bin':(source/'Native/NewtonRejected.00054.bin').read_bytes(),
 'Input/ActualRefinedNewton.bin':(source/'Native/NewtonRejected.00055.bin').read_bytes(),
 'SourceRecordedTailSystemProof.json':(parent/'RecordedTailSystemProof.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ActualTailNewtonFrozenStressScan',actual_import_this_pass=['modTestCoupledElasticState'],
 actual_imports_this_replay=[],actual_component_count=257,immutable_compiled_book_reused=True,compiled_source_root=parent.name,
 exact_history25_and_recorded54_current=True,base_system_reproduced_in925_before_scan=True,
 no_new_NDS_search_or_material_history_acceptance=True,endpoint_fractions=fractions,
 physical_capacity_certified=False,no_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'EXACT COMPILED925;18 SIGNED LOCAL ENDPOINTS;NO NEW NDS OR HISTORY ACCEPTANCE')
