"""ДЛЯ ТЕСТОВ. Проверяет фактический код исправленного частного продолжения ASD."""
from pathlib import Path

here=Path(__file__).resolve().parent;source=here/'Check-TestASDGeometricSource.py';code=source.read_text('utf-8-sig')
code=code.replace("code=code.replace('ProbeBudgetSourceProof.json','GeometricSourceProof.json')",
    "code=code.replace('Cold1023ASDGeometricContinuation_20261126','Cold1024ASDGeometricDecline_20261126')\ncode=code.replace('ProbeBudgetSourceProof.json','GeometricSourceProof.json')")
exec(compile(code,str(source),'exec'),dict(__file__=str(Path(__file__).resolve()),__name__='ASD_geometric_decline_source_test'))
