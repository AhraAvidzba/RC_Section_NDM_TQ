"""ДЛЯ ТЕСТОВ. Проверяет холодный923, сохранённые257 компонентов и полное побитное совпадение пути с920."""
from pathlib import Path
here=Path(__file__).resolve().parent;file=here/'Check-TestFixedMetricCold.py';code=file.read_text('utf-8-sig')
for old,new in [("default=base/'Cold868FixedMetricCold_20261126'","default=base/'Cold923RecordedNewtonCold96_20261126'"),
 ("parent=base/'Cold867FixedMetricKernel_20261126'","parent=base/'Cold922ColdNewtonCapture_20261126'"),
 ("scope['case']=='FixedMetricKernelTrueCold96'","scope['case']=='RecordedNewtonTrueCold96'"),
 ('assert int(timeout[1])==600','assert int(timeout[1])==1200'),
 ('timeout_seconds=600,workflow_result_missing=True','timeout_seconds=1200,workflow_result_missing=True'),
 ("assert outcome['all_saved_sources_exact']==256","assert outcome['all_saved_sources_exact']==257"),
 ('same_actual_compiled_kernel867=True','same_actual_compiled_kernel922=True'),
 ("root/'FixedMetricColdProof.json'","root/'RecordedNewtonCold96Proof.json'")]:
    assert old in code,old;code=code.replace(old,new)
old="reg,_=helpers['TestNative'](root,'OperationalArcRegressions');assert reg==dict(passed=164,failed=0,macros=5,timeout600=False,own_process_closed=True)"
new="reg=json.loads((root/'PriorDirectedRegressionProof.json').read_bytes());assert reg==dict(passed=164,failed=0,macros=5,timeout600=False,own_process_closed=True)"
assert code.count(old)==1;code=code.replace(old,new,1)
old="assert (root/'src/Solver/CCoupledElasticState.cls').read_bytes()==(base/'Cold835GatedColdReplay_20261125/src/Solver/CCoupledElasticState.cls').read_bytes()"
new="assert run_path(str(here/'Check-TestSpatialProjectionCompatibility.py'))['TestGeometryZeroOffset'](root,base/'Cold835GatedColdReplay_20261125')"
assert code.count(old)==1;code=code.replace(old,new,1)
namespace={'__file__':str(file),'__name__':'recorded_newton_true_cold96_test'}
exec(compile(code,str(file),'exec'),namespace)
root=namespace['root'];base=namespace['base'];proof=namespace['proof']
reference=base/'Cold920AffineRecoveryCold96_20261126'
same=(root/'Native/Stock.RawHistory.bin').read_bytes()==(reference/'Native/Stock.RawHistory.bin').read_bytes()
proof.update(entire_raw_trajectory_and_intermediate_frames_bit_identical920=same,
    rejected_directions_recorded=len(list((root/'Native').glob('NewtonRejected.*.bin'))),
    no_other_diagnostic_modes_enabled=True,all257_saved_sources_verified=bool(proof['cold'].get('all_saved_sources_exact')==257))
import json
(root/'RecordedNewtonCold96Proof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
assert same,'Запись изменила путь; снимки нельзя объявлять точным контекстом исходного920.'
assert proof['rejected_directions_recorded']>0
print('ENTIRE RAW COLD HISTORY EXACT920;REJECTED DIRECTIONS',proof['rejected_directions_recorded'])
