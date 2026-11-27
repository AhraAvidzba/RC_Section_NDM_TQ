"""ДЛЯ ТЕСТОВ. Повторяет аудит рабочих классов и проверяет единственное тестовое подключение изменяющего копии наблюдателя."""
from pathlib import Path

here = Path(__file__).resolve().parent
source = here / 'Check-TestWorkingASDSource.py'
code = source.read_text('utf-8-sig')
for before, after in [
    ('Cold1067WorkingASDNoObserver_20261127', 'Cold1068WorkingASDObserverMutation_20261127'),
    ("compiled['Modules'] == classes", "compiled['Modules'] == ['modTestSelectableMaterialCore']"),
    ('if name not in classes:', "if name != 'modTestSelectableMaterialCore':"),
    ('all252_other_components_unchanged=True', 'all255_other_components_unchanged=True'),
    ('ACTUAL WORKING ASD SOURCE VERIFIED;4 CLASSES EXACT WORKTREE;252 OTHER COMPONENTS UNCHANGED;NO TEST EXPORTS',
     'WORKING ASD OBSERVER TEST SOURCE VERIFIED;4 NUMERICAL CLASSES UNCHANGED;ONLY TEST CONNECTION ADDED')
]:
    assert code.count(before) == 1, before
    code = code.replace(before, after, 1)
anchor = 'proof = dict('
assert code.count(anchor) == 1
code = code.replace(anchor, '''assert 'CCoupledProgressObserverTest' in new
assert normal(body((repo / 'tests/CCoupledProgressObserverTest.cls').read_text('utf-8-sig'), 'CCoupledProgressObserverTest')[0])[0] == normal(body(new['CCoupledProgressObserverTest'], 'CCoupledProgressObserverTest')[0])[0]
staged_test = (root / 'tests/modTestSelectableMaterialCore.bas').read_text('utf-8-sig')
assert normal(body(staged_test, 'modTestSelectableMaterialCore')[0])[0] == normal(body(new['modTestSelectableMaterialCore'], 'modTestSelectableMaterialCore')[0])[0]
''' + anchor, 1)
exec(compile(code, str(source), 'exec'), dict(__file__=str(Path(__file__).resolve()), __name__='TestCheckWorkingObserver'))
