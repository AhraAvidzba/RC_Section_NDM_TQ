"""ДЛЯ ТЕСТОВ. Сверяет драйвер чистого кручения с сохранением всех компонентов ядра."""
from pathlib import Path

here=Path(__file__).resolve().parent;source=here/'Check-TestASDCapacityScalingSource.py';code=source.read_text('utf-8-sig')
code=code.replace('Cold1027ASDCapacityScaling_20261126','Cold1035ASDPureTorsion_20261126')
exec(compile(code,str(source),'exec'),dict(__file__=str(Path(__file__).resolve()),__name__='ASD_pure_torsion_source_test'))
