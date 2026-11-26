"""ДЛЯ ТЕСТОВ. Проверяет изменённый переключатель QR-резерва чистого N и неизменность 255 остальных компонентов книги."""
from pathlib import Path

here = Path(__file__).resolve().parent
source = here / 'Check-TestASDHybridSource.py'
code = source.read_text('utf-8-sig')
assert code.count('Cold1033ASDHybridRecovery_20261126') == 1
code = code.replace('Cold1033ASDHybridRecovery_20261126', 'Cold1041ASDPureNHybrid_20261127')
exec(compile(code, str(source), 'exec'), dict(__file__=str(Path(__file__).resolve()), __name__='TestAxialHybridSource'))
