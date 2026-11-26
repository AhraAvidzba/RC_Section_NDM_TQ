"""ДЛЯ ТЕСТОВ. Повторно использует строгий аудит сохранённой книги для одного читающего наблюдателя точной нагрузки."""
from pathlib import Path

here = Path(__file__).resolve().parent
source = here / 'Check-TestASDDirectRecoverySource.py'
code = source.read_text('utf-8-sig')
for before, after in [
    ('Cold1062ASDPureNDirectRecovery_20261127', 'Cold1063ASDPureNDirectObservation_20261127'),
    ("['CCoupledLoadContinuation', 'CCoupledSectionAnalysis']", "['CCoupledLoadContinuation']"),
    ('all254_other_components_unchanged=True', 'all255_other_components_unchanged=True'),
    ('new_direct_recovery_ASD_only=True', 'only_reading_observations_added=True'),
    ('DIRECT ASD RECOVERY SOURCE VERIFIED;ONLY TWO EXPECTED COMPONENTS;MATERIALS AND TOLERANCES UNCHANGED',
     'DIRECT ASD OBSERVATION VERIFIED;ONLY ONE READ-ONLY COMPONENT CHANGE')
]:
    assert code.count(before) == 1, before
    code = code.replace(before, after, 1)
start = code.index("patch = run_path(str(here / 'Patch-TestASDDirectRecovery.py'))")
end = code.index('for name in compiled[', start)
code = code[:start] + '''patch = run_path(str(here / 'Patch-TestDirectRecoveryObservation.py'))
name = 'CCoupledLoadContinuation'
expected = patch['TestPatchDirectRecoveryObservation'](body(old[name], name)[0], canonical=True)
assert normal(expected)[0] == normal(body(new[name], name)[0])[0]
''' + code[end:]
exec(compile(code, str(source), 'exec'), dict(__file__=str(Path(__file__).resolve()), __name__='TestCheckDirectObservation'))
