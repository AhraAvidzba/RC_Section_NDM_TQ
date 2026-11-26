"""ДЛЯ ТЕСТОВ. Проверяет действительный холодный исход Newton901 и сохранённую повторно открытую книгу без объявления NumFail разрушением."""
from pathlib import Path
here=Path(__file__).resolve().parent;file=here/'Check-TestFreshArcCold96.py'
code=file.read_text('utf-8-sig')
for old,new in [('Cold895FreshSpatialArc_20261126','Cold901ConsistentControlRefinement_20261126'),
                ('Cold897FreshArcCold96_20261126','Cold904ConsistentRefinementCold96_20261126'),
                ('FreshArcTrueCold96','ConsistentRefinementTrueCold96'),('same_actual_compiled_kernel895','same_actual_compiled_kernel901'),
                ('FreshArcCold96Proof.json','ConsistentRefinementCold96Proof.json')]:
    assert old in code,old;code=code.replace(old,new)
exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'consistent_refinement_cold96_test'})
