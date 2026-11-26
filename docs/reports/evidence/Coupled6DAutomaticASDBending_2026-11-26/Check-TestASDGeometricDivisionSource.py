"""ДЛЯ ТЕСТОВ. Проверяет точный частный контроллер деления шага ASD и сохранность остальных компонентов."""
from pathlib import Path

here=Path(__file__).resolve().parent;source=here/'Check-TestASDGeometricSource.py'
code=source.read_text('utf-8-sig').replace('Cold1023ASDGeometricContinuation_20261126','Cold1025ASDGeometricDivision_20261126')
exec(compile(code,str(source),'exec'),dict(__file__=str(Path(__file__).resolve()),__name__='ASD_geometric_division_source_test'))
