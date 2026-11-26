"""ДЛЯ ТЕСТОВ. Читает все локальные напряжения и поверхности вдоль секущей и фактического рабочего касательного прогноза холодного32."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold900Cold32OutgoingJacobian_20261126';root=base/'Cold906Cold32OutgoingStressScan_20261126'
proof=json.loads((parent/'Cold32OutgoingJacobianProof.json').read_bytes())
assert proof['outgoing_working_predictor']['relative_difference_from_independent_linear_solve']<1e-8
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
r=R((parent/'Native/Cold32DirectionalJacobian.bin').read_bytes())
meta=r.VectorTest();accepted=r.VectorTest();movement=r.VectorTest();projection=r.VectorTest()
assert r.VectorTest().tolist()==[1.]
tangent=r.VectorTest();probe=r.VectorTest();np.testing.assert_array_equal(probe,accepted+movement/1024)
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](parent/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)])
M=np.einsum('spcd,spce,p,c->de',op,op,coords[:,2],np.array([1,1,1,.5,.5,.5]))/(3*sum(coords[:,2]));M[56:]=0;M[:,56:]=0
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Input/Reverse1.bin')
alpha=.00010929902223202709;dl=meta[0]-meta[1]
orientation=alpha*alpha*dl+float(projection@tangent)*meta[6]
scale=float(np.sign(orientation)*meta[6]/np.sqrt(tangent@M@tangent));outgoing=scale*tangent
predictions=[accepted,accepted+movement,accepted+outgoing]
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
prep=json.loads((root/'Preparation.json').read_bytes());name='tools/build_workbook/Run-Audit02ModeTests.ps1'
runner=(base/'Cold857SpatialDeclineRecovery_20261125'/name).read_text('utf-8-sig')
anchor='        else { $result = [string]$excel.Run($macroName) }';assert runner.count(anchor)==1
runner=runner.replace(anchor,anchor+'''
        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected cold32 stress-scan setup macro.'}
        $null=$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestReadUnresolvedArcScan")
        $lines.Add('READONLY_PRIVATE_SCAN_COMPLETED: no new equilibrium search')
        Save-Progress
''',1)
cases=[dict(case=1,name='zero',lambda_increment=0),dict(case=2,name='incoming_secant',lambda_increment=float(dl)),
       dict(case=3,name='actual_outgoing_tangent',lambda_increment=scale)]
values={name:runner.replace('\n','\r\n').encode('utf-8-sig'),
 'Input/ArcScanFields.bin':vector([3])+b''.join(vector(f) for f in predictions),
 'Input/ArcScanCases.json':(json.dumps(cases,sort_keys=True,indent=2)+'\n').encode(),
 'SourceOutgoingJacobianProof.json':(parent/'Cold32OutgoingJacobianProof.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='Cold32ActualOutgoingLocalStressScan',actual_import_this_pass=json.loads((parent/'docs/regression/CodeUpdate.json').read_bytes())['Modules'],
 actual_imports_this_replay=[],actual_component_count=256,immutable_compiled_book_reused=True,compiled_source_root=parent.name,
 actual_history32_exact=True,actual_operational_tangent_from900=True,alpha=alpha,outgoing_scale=scale,
 three_segments_eight_levels_all129_concrete_stations=True,no_new_NDS_search_or_material_acceptance=True,
 physical_capacity_certified=False,no_working_or_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
print(root,'EXACT COMPILED900;READONLY POINT STRESSES/SURFACES;NO NDS SEARCH;SCALE',scale)
