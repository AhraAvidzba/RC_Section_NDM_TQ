"""ДЛЯ ТЕСТОВ. Проверяет согласованность направлений с пластической стороной фактического перехода точки11."""
from pathlib import Path
here=Path(__file__).resolve().parent;file=here/'Check-TestSurfaceDirectionSigns.py'
code=file.read_text('utf-8-sig')
for old,new in [('Check-TestSurfaceMatrices.py','Check-TestPoint11Matrices.py'),
 ('point=5,side=1','point=11,side=-1'),('TestYield(5,1,','TestYield(11,-1,')]:
    assert old in code,old;code=code.replace(old,new)
old="scope=run_path(str(here/'Check-TestPoint11Matrices.py'))"
assert old in code;code=code.replace(old,old+"['scope']",1)
scope={'__file__':str(file),'__name__':'point11_direction_signs_test'}
exec(compile(code,str(file),'exec'),scope)
