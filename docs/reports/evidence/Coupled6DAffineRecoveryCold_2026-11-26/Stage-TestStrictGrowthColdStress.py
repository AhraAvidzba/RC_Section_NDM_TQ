"""ДЛЯ ТЕСТОВ. Готовит только материальный повтор завершённой холодной истории890 с точным кодом моделей и сборки881."""
from pathlib import Path
import sys
here=Path(__file__).resolve().parent;file=here/'Stage-TestImmutableColdStressReplay.py';code=file.read_text('utf-8-sig')
assert code.count("compiled=base/'Cold855BalancedInitialControl_20261125'")==1
code=code.replace("compiled=base/'Cold855BalancedInitialControl_20261125'","compiled=base/'Cold881SpatialSteelLimits_20261126'",1)
code=code.replace("if partial:assert cold['timeout_seconds']==600","if partial:assert cold['timeout_seconds']==1200",1)
sys.argv=[str(file),'--source','Cold890StrictGrowthCold96_20261126','--root','Cold892StrictGrowthColdStress_20261126','--proof','StrictGrowthCold96Proof.json']
exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'strict_growth_cold_stress_staging'})
