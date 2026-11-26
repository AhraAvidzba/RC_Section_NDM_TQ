"""ДЛЯ ТЕСТОВ. Готовит материальный повтор завершённого холодного897 без нового поиска НДС или предела."""
from pathlib import Path
here=Path(__file__).resolve().parent
file=here/'Stage-TestStrictGrowthColdStress.py'
code=file.read_text('utf-8-sig')
for old,new in [('Cold890StrictGrowthCold96_20261126','Cold897FreshArcCold96_20261126'),
                ('Cold892StrictGrowthColdStress_20261126','Cold898FreshArcColdStress_20261126'),
                ('StrictGrowthCold96Proof.json','FreshArcCold96Proof.json')]:
    assert old in code
    code=code.replace(old,new)
exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'fresh_arc_cold_stress_staging'})
