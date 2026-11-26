"""ДЛЯ ТЕСТОВ. Запускает контроллер новой дуговой пробы с нуля после положительных проверок контекста и164 регрессий."""
from pathlib import Path
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
proof=json.loads((base/'Cold895FreshSpatialArc_20261126/FreshSpatialArcProof.json').read_bytes())
assert proof['old_exact_retained_arc_gate_unchanged'] and proof['native']['passed']==23
file=here/'Stage-TestStrictGrowthCold96.py';code=file.read_text('utf-8-sig')
code=code.replace('Cold888StrictEnvelopeGrowth_20261126','Cold895FreshSpatialArc_20261126').replace('Cold890StrictGrowthCold96_20261126','Cold897FreshArcCold96_20261126').replace('Cold889StrictGrowthRegressions_20261126','Cold896FreshArcRegressions_20261126')
code=code.replace("audit['imported_this_pass']==['CCoupledDeformationContinuation']","audit['imported_this_pass']==['CCoupledDeformationContinuation','modTestCoupledElasticState']")
code=code.replace('StrictGrowthTrueCold96','FreshArcTrueCold96').replace('exact_compiled_source888','exact_compiled_source895')
code=code.replace('ACTUAL COLD;EXACT888;','ACTUAL COLD;EXACT895;')
anchor="assert 'If mProbeCount >= PROBE_LIMIT Then' in controller"
assert code.count(anchor)==1
code=code.replace(anchor,anchor+"\nassert controller.count('preferArc = CanStartFreshSpatialArc(history, current)')==1\nassert controller.count('Private Function CanStartFreshSpatialArc(')==1",1)
exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'fresh_arc_cold96_staging'})
