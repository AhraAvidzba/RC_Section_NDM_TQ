"""ДЛЯ ТЕСТОВ. Проверяет материальный повтор холодной истории897 на точных моделях и кинематике881."""
from pathlib import Path
here=Path(__file__).resolve().parent
file=here/'Check-TestStrictGrowthColdStress.py'
code=file.read_text('utf-8-sig')
for old,new in [('Cold890StrictGrowthCold96_20261126','Cold897FreshArcCold96_20261126'),
                ('Cold892StrictGrowthColdStress_20261126','Cold898FreshArcColdStress_20261126')]:
    assert old in code
    code=code.replace(old,new)
exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'__main__'})
