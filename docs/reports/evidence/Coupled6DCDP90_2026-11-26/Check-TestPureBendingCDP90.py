"""ДЛЯ ТЕСТОВ. Проверяет состояния отдельного CDP-прогона с разрешённым лимитом 5400 с при неизменных физических критериях."""
from pathlib import Path

here=Path(__file__).resolve().parent
source=here/'Check-TestPureBendingCDPLong.py'
code=source.read_text('utf-8-sig')
assert 'Cold1014PureBendingCDPLong_20261126' in code and '==1800' in code
code=code.replace('Cold1014PureBendingCDPLong_20261126','Cold1019PureBendingCDP90_20261126')
code=code.replace('1800','5400').replace('long_CDP_curve_check_test','CDP_90min_curve_check_test')
exec(compile(code,str(source),'exec'),dict(__file__=str(Path(__file__).resolve()),__name__='CDP_90min_check_test'))
