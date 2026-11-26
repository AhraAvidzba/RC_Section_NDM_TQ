"""ДЛЯ ТЕСТОВ. Готовит те же восемь дуговых проб с исходным Arc901; новый читающий метод Newton остаётся невызываемым."""
from pathlib import Path
import json
here=Path(__file__).resolve().parent;file=here/'Stage-TestResidualArcPredictor.py';code=file.read_text('utf-8-sig')
code=code.replace('Cold908ResidualArcPredictor_20261126','Cold909ResidualArcPredictorBaseline_20261126')
anchor="values[name]=code.replace('\\n','\\r\\n').encode('utf-8-sig')\ntree=ast.parse"
replacement="""code=(parent/'src/Solver/CCoupledArcContinuation.cls').read_text('utf-8-sig')+'''

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает отсутствие нового сравнения в исходном алгоритме дуги; материал и принятая история не меняются.
Friend Property Get TestPredictorComparison() As Variant
    TestPredictorComparison = Empty
End Property
'''
"""+anchor
assert code.count(anchor)==1;code=code.replace(anchor,replacement,1)
exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'residual_predictor_baseline_staging'})
root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold909ResidualArcPredictorBaseline_20261126');prep=json.loads((root/'Preparation.json').read_bytes())
prep['directed_scope'].update(case='ResidualArcPredictorOriginalBaseline',baseline_arc_step_exact901=True,
    predictor_selection_not_called=True,extra_three_readonly_residual_evaluations_inside_one_predictor_preparation=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
