"""ДЛЯ ТЕСТОВ. Проверяет весь материальный повтор холодного890 при неизменных моделях и методах881, без нового поиска НДС."""
from pathlib import Path
here=Path(__file__).resolve().parent;file=here/'Check-TestCriticalHColdStress.py';code=file.read_text('utf-8-sig')
code=code.replace('Cold882CriticalHColdStress_20261126','Cold892StrictGrowthColdStress_20261126').replace('Cold880CriticalHCold96_20261126','Cold890StrictGrowthCold96_20261126').replace('Cold855BalancedInitialControl_20261125','Cold881SpatialSteelLimits_20261126')
exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'__main__'})
