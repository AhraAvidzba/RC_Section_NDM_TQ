"""ДЛЯ ТЕСТОВ. Проверяет обычное деление неудачного геометрического шага ASD до попытки смены материальной меры."""
from pathlib import Path

here=Path(__file__).resolve().parent;source=here/'Stage-TestASDGeometricDecline.py';code=source.read_text('utf-8-sig')
code=code.replace('Cold1024ASDGeometricDecline_20261126','Cold1025ASDGeometricDivision_20261126').replace('PREPARED1024','PREPARED1025')
needle="code=code.replace(needle,injection+needle,1)";assert code.count(needle)==1
code=code.replace(needle,"""injection += \"assert code.count('If Not unresolvedSteel Then')==1\\ncode=code.replace('If Not unresolvedSteel Then','If Not unresolvedSteel And Not UsesASDGeometricControl() Then',1)\\n\"
"""+needle,1)
exec(compile(code,str(source),'exec'),dict(__file__=str(Path(__file__).resolve()),__name__='ASD_geometric_division_stage_test'))
