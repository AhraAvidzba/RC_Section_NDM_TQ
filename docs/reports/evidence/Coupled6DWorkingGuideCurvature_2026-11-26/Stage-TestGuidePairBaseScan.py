"""ДЛЯ ТЕСТОВ. Повторно читает материал в принятой исходной точке и действительных полной/половинной пробах949 без решения новых НДС."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold949WorkingGuidePair_20261126';root=base/'Cold950GuidePairBaseScan_20261126'
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((parent/'Native/WorkingGuidePairs.bin').read_bytes())
assert r.raw[:16]==b'NDMGUIDEPAIR0001' and r.VectorTest().tolist()==[1.]
info,accepted,full,half,projection=[r.VectorTest() for _ in range(5)]
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
cases=[dict(case=1,name='accepted25_read_again',lambda_value=float(info[2])),
    dict(case=2,name='actual_full949',lambda_value=float(info[0])),dict(case=3,name='actual_half949',lambda_value=float(info[1]))]
updates={'Input/ArcScanFields.bin':vector([3.])+b''.join(vector(v) for v in (accepted,full,half)),
    'Input/ArcScanCases.json':json.dumps(cases,sort_keys=True,indent=2).encode(),
    'SourceWorkingGuidePairs.bin':(parent/'Native/WorkingGuidePairs.bin').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(parent/name).read_text('utf-8-sig')
old="        $lines.Add([string]$excel.Run(\"'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestRunWorkingGuidePair\"))"
assert runner.count(old)==1
new="        $null=$excel.Run(\"'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestReadUnresolvedArcScan\")\n        $lines.Add('READONLY_PRIVATE_SCAN_COMPLETED: no new equilibrium search')"
updates[name]=runner.replace(old,new,1).replace('\n','\r\n').encode('utf-8-sig')
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='Accepted25AndActualFullHalfMaterialReread',immutable_compiled949=True,
    no_new_compilation_or_NDS_solve=True,material_history25_unchanged=True,no_acceptance=True,
    source949_fields_used_exactly=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'READONLY ACCEPTED/FULL/HALF;SAME MATERIAL HISTORY25')
