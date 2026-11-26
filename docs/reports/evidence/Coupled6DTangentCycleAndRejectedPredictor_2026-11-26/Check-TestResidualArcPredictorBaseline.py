"""ДЛЯ ТЕСТОВ. Проверяет тот же набор дуговых проб при исходном алгоритме Arc901 и отдельном пустом тестовом чтении."""
from pathlib import Path
here=Path(__file__).resolve().parent;file=here/'Check-TestResidualArcPredictor.py';code=file.read_text('utf-8-sig')
code=code.replace('Cold908ResidualArcPredictor_20261126','Cold909ResidualArcPredictorBaseline_20261126')
code=code.replace('ResidualArcPredictorProof.json','ResidualArcPredictorBaselineProof.json')
code=code.replace('only_predictor_comparison_and_selection_added=True','baseline_arc_step_exact901=True')
anchor="assert set(right)==set(left)|{'TestPredictorComparison'}"
assert code.count(anchor)==1;code=code.replace(anchor,anchor+"\nassert left['TryStep']==right['TryStep']",1)
exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'residual_predictor_baseline_test'})
