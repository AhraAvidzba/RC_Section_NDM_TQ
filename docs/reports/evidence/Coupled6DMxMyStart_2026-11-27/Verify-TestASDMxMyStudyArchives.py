"""ДЛЯ ТЕСТОВ: сверяет все прямые и унаследованные файлы исследования начала Mx/My и точного отказа."""
from pathlib import Path

source=Path(__file__).with_name('Verify-TestASDRealisticArchives.py');code=source.read_text('utf-8-sig')
before="root=Path.cwd()/'docs/reports/evidence/Coupled6DASDRealisticSection_2026-11-27'"
after="root=Path.cwd()/'docs/reports/evidence/Coupled6DMxMyStart_2026-11-27'"
assert code.count(before)==1;code=code.replace(before,after,1)
before="for part in ('Bridge','Results'):";assert code.count(before)==1
code=code.replace(before,"for part in ('Frozen','Start'):",1)
exec(compile(code,str(source),'exec'),{'__file__':str(source),'__name__':'__main__'})
