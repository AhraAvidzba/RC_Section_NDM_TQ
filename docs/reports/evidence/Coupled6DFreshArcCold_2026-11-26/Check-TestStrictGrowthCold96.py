"""ДЛЯ ТЕСТОВ. Проверяет исход штатного холодного запуска строгого адаптивного контроллера и фактически сохранённую книгу."""
from pathlib import Path
here=Path(__file__).resolve().parent
file=here/'Check-TestCriticalHCold96.py';code=file.read_text('utf-8-sig')
code=code.replace('Cold880CriticalHCold96_20261126','Cold890StrictGrowthCold96_20261126').replace('Cold878CriticalHRecoveryPairs_20261126','Cold888StrictEnvelopeGrowth_20261126')
code=code.replace('CriticalHTrueCold96','StrictGrowthTrueCold96').replace('same_actual_compiled_kernel878','same_actual_compiled_kernel888').replace('CriticalHCold96Proof.json','StrictGrowthCold96Proof.json')
anchor="exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'critical_H_true_cold96_test'})"
assert code.count(anchor)==1
old="assert (root/'src/Solver/CCoupledElasticState.cls').read_bytes()==(base/'Cold835GatedColdReplay_20261125/src/Solver/CCoupledElasticState.cls').read_bytes()"
new="assert run_path(str(here/'Check-TestSpatialProjectionCompatibility.py'))['TestGeometryZeroOffset'](root,base/'Cold835GatedColdReplay_20261125')"
code=code.replace(anchor,'code=code.replace('+repr(old)+','+repr(new)+')\n'+anchor,1)
exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'strict_growth_cold96_test'})
