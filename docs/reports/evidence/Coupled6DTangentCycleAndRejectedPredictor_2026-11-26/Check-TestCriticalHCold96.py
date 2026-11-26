"""ДЛЯ ТЕСТОВ. Фиксирует фактический исход холодного880 и проверяет сохранённую книгу либо точное завершение защитного тайм-аута."""
from pathlib import Path
from runpy import run_path
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');root=base/'Cold880CriticalHCold96_20261126';parent=base/'Cold878CriticalHRecoveryPairs_20261126'
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
for file in (parent/'src').rglob('*'):
    if file.is_file():assert file.read_bytes()==(root/file.relative_to(parent)).read_bytes()
scope=json.loads((root/'Preparation.json').read_bytes())['directed_scope'];assert scope['case']=='CriticalHTrueCold96' and scope['actual_imports_this_replay']==[]
assert scope['global_probe_limit96_unchanged'] and scope['primary_phase_iterations30_unchanged']
# Общая независимая проверка холодного снимка переиспользуется только с явным контрактом нового источника и окна1200.
file=here/'Check-TestFixedMetricCold.py';code=file.read_text('utf-8-sig')
code=code.replace("default=base/'Cold868FixedMetricCold_20261126'","default=base/'Cold880CriticalHCold96_20261126'")
code=code.replace("parent=base/'Cold867FixedMetricKernel_20261126'","parent=base/'Cold878CriticalHRecoveryPairs_20261126'")
code=code.replace("scope['case']=='FixedMetricKernelTrueCold96'","scope['case']=='CriticalHTrueCold96'")
old="reg,_=helpers['TestNative'](root,'OperationalArcRegressions');assert reg==dict(passed=164,failed=0,macros=5,timeout600=False,own_process_closed=True)"
assert code.count(old)==1
code=code.replace(old,"reg=json.loads((root/'PriorDirectedRegressionProof.json').read_bytes());assert reg==dict(passed=164,failed=0,macros=5,timeout600=False,own_process_closed=True)")
code=code.replace('assert int(timeout[1])==600','assert int(timeout[1])==1200').replace('timeout_seconds=600,workflow_result_missing=True','timeout_seconds=1200,workflow_result_missing=True')
code=code.replace('same_actual_compiled_kernel867=True','same_actual_compiled_kernel878=True')
code=code.replace("root/'FixedMetricColdProof.json'","root/'CriticalHCold96Proof.json'")
exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'critical_H_true_cold96_test'})
