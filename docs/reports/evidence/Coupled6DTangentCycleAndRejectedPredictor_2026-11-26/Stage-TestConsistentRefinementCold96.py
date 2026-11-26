"""ДЛЯ ТЕСТОВ. Запускает единственное исправление Newton с нуля после164 регрессий; холодный путь не получает исследовательских полей."""
from pathlib import Path
here=Path(__file__).resolve().parent;file=here/'Stage-TestFreshArcCold96.py'
code=file.read_text('utf-8-sig')
for old,new in [('Cold895FreshSpatialArc_20261126','Cold901ConsistentControlRefinement_20261126'),
                ('Cold897FreshArcCold96_20261126','Cold904ConsistentRefinementCold96_20261126'),
                ('Cold896FreshArcRegressions_20261126','Cold903ConsistentRefinementRegressions_20261126'),
                ("['CCoupledDeformationContinuation','modTestCoupledElasticState']","['CCoupledNewtonSolver','modTestCoupledElasticState']"),
                ('FreshArcTrueCold96','ConsistentRefinementTrueCold96'),('exact_compiled_source895','exact_compiled_source901'),
                ('ACTUAL COLD;EXACT895;','ACTUAL COLD;EXACT901;')]:
    assert old in code,old;code=code.replace(old,new)
exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'consistent_refinement_cold96_staging'})
