"""ДЛЯ ТЕСТОВ. Сохраняет отрицательную пробу адаптации и включает принятие непроверенного спада для выбранной геометрической меры ASD."""
from pathlib import Path

here=Path(__file__).resolve().parent;source=here/'Stage-TestASDGeometricContinuation.py';code=source.read_text('utf-8-sig')
code=code.replace('Cold1023ASDGeometricContinuation_20261126','Cold1024ASDGeometricDecline_20261126').replace('PREPARED1023','PREPARED1024')
needle="(root/name).write_bytes(code.replace('\\n','\\r\\n').encode('utf-8-sig'))"
assert code.count(needle)==1
injection="""needle='reinforced And (mUsesTransverseControl Or mUsesFieldControl)'
assert code.count(needle)==2
code=code.replace(needle,'reinforced And (mUsesTransverseControl Or mUsesFieldControl Or UsesASDGeometricControl())')
"""
code=code.replace(needle,injection+needle,1)
exec(compile(code,str(source),'exec'),dict(__file__=str(Path(__file__).resolve()),__name__='ASD_geometric_decline_stage_test'))
