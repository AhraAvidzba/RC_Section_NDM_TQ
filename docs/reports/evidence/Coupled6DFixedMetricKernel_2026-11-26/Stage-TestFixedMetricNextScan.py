"""ДЛЯ ТЕСТОВ. Читает отклик нового проверенного состояния20 и фактического неудачного касательного прогноза без изменения VBA."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold867FixedMetricKernel_20261126';root=base/'Cold869FixedMetricNextScan_20261126'
proof=json.loads((parent/'FixedMetricKernelProof.json').read_bytes());assert proof['new_accepted_equilibria']==1 and proof['all_new_H_strictly_growing']
read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'];source=read(parent/'Input/Reverse1.bin')
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=Reader((parent/'Native/LoadingWalk.002.bin').read_bytes())
packet=[r.VectorTest() for _ in range(15)];full=r.VectorTest();actual=r.VectorTest();projection=r.VectorTest()
factor=packet[0][23];assert factor==proof['attempts'][1]['lambda_value']
frame=dict(fields=np.vstack([source['fields'],full]),fractions=np.r_[source['fractions'],1.],lambdas=np.r_[source['lambdas'],factor])
g=Reader((parent/'Native/LoadingWalk.Guide003.bin').read_bytes());g.VectorTest()
for size in (3,3):
    if g.VectorTest()[0]:
        for _ in range(size):g.VectorTest()
assert g.VectorTest()[0]==1
tm,probe,perlambda,movement=[g.VectorTest() for _ in range(4)];assert g.position==len(g.raw)
predictions=[full,full+.125*(full-source['fields'][-1]),full+proof['attempts'][-1]['ratio']*movement]
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
prep=json.loads((root/'Preparation.json').read_bytes());vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
script=(base/'Cold857SpatialDeclineRecovery_20261125/tools/build_workbook/Run-Audit02ModeTests.ps1').read_text('utf-8-sig')
anchor='        else { $result = [string]$excel.Run($macroName) }';assert script.count(anchor)==1
script=script.replace(anchor,anchor+'''
        # ДЛЯ ТЕСТОВ: только читающий контроль принятого состояния20 и отрезков прогноза.
        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected next-state setup macro.'}
        $null = $excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestReadUnresolvedArcScan")
        $lines.Add('READONLY_PRIVATE_SCAN_COMPLETED: no new equilibrium search')
        Save-Progress
''',1)
values={'Input/Reverse1.bin':run_path(str(here/'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](source,frame),
 'Input/ArcScanFields.bin':vector([3])+b''.join(vector(f) for f in predictions),
 'tools/build_workbook/Run-Audit02ModeTests.ps1':script.replace('\n','\r\n').encode('utf-8-sig'),
 'SourceVerifiedWarm20Proof.json':(parent/'FixedMetricKernelProof.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='VerifiedWarm20ZeroAndOutgoingStressScan',actual_import_this_pass=json.loads((parent/'docs/regression/CodeUpdate.json').read_bytes())['Modules'],
 actual_imports_this_replay=[],compiled_source_root=parent.name,source_states=20,readonly_queries_only=True,
 predictions=['exact accepted20','1/8 incoming movement','last failed ratio times actual outgoing tangent movement'],no_new_NDS_or_acceptance=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
print(root,'SAME COMPILED867;TRUE VERIFIED20;READONLY ZERO/INCOMING/ACTUAL FAILED OUTGOING SCAN')
