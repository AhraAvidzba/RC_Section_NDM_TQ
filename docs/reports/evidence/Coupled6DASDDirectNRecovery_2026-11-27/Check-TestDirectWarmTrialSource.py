"""ДЛЯ ТЕСТОВ. Применяет строгий аудит одного изменённого компонента к порядку точных силовых проб ASD."""
from pathlib import Path

here = Path(__file__).resolve().parent
source = here / 'Check-TestASDDirectRecoverySource.py'
code = source.read_text('utf-8-sig')
for before, after in [
    ('Cold1062ASDPureNDirectRecovery_20261127', 'Cold1064ASDPureNDirectWarmTrial_20261127'),
    ("['CCoupledLoadContinuation', 'CCoupledSectionAnalysis']", "['CCoupledLoadContinuation']"),
    ('all254_other_components_unchanged=True', 'all255_other_components_unchanged=True'),
    ('new_direct_recovery_ASD_only=True', 'unseeded_exact_trial_before_interpolated_trial=True'),
    ('DIRECT ASD RECOVERY SOURCE VERIFIED;ONLY TWO EXPECTED COMPONENTS;MATERIALS AND TOLERANCES UNCHANGED',
     'DIRECT ASD WARM EXACT TRIAL SOURCE VERIFIED;ONLY EXPECTED ONE-COMPONENT CHANGE')
]:
    assert code.count(before) == 1, before
    code = code.replace(before, after, 1)
start = code.index("patch = run_path(str(here / 'Patch-TestASDDirectRecovery.py'))")
end = code.index('for name in compiled[', start)
code = code[:start] + '''patch = run_path(str(here / 'Patch-TestDirectWarmTrial.py'))
name = 'CCoupledLoadContinuation'
expected = patch['TestPatchDirectWarmTrial'](body(old[name], name)[0], canonical=True)
assert normal(expected)[0] == normal(body(new[name], name)[0])[0]
''' + code[end:]
exec(compile(code, str(source), 'exec'), dict(__file__=str(Path(__file__).resolve()), __name__='TestCheckDirectWarmTrial'))
