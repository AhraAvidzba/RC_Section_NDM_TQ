"""ДЛЯ ТЕСТОВ. Проверяет штатный холодный поиск нового дугового контекста, фактически сохранённую книгу и физические масштабы."""
from pathlib import Path
here=Path(__file__).resolve().parent;file=here/'Check-TestStrictGrowthCold96.py';code=file.read_text('utf-8-sig')
code=code.replace('Cold890StrictGrowthCold96_20261126','Cold897FreshArcCold96_20261126').replace('Cold888StrictEnvelopeGrowth_20261126','Cold895FreshSpatialArc_20261126')
code=code.replace('StrictGrowthTrueCold96','FreshArcTrueCold96').replace('same_actual_compiled_kernel888','same_actual_compiled_kernel895').replace('StrictGrowthCold96Proof.json','FreshArcCold96Proof.json')
exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'fresh_arc_cold96_test'})
