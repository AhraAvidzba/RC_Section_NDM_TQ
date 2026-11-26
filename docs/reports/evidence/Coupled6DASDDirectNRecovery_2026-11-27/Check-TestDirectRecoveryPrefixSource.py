"""ДЛЯ ТЕСТОВ. Проверяет точную правку раннего осевого префикса ASD в фактически собранной книге."""
from pathlib import Path

here = Path(__file__).resolve().parent
source = here / 'Check-TestASDDirectRecoverySource.py'
code = source.read_text('utf-8-sig')
for before, after in [
    ('Cold1062ASDPureNDirectRecovery_20261127', 'Cold1065ASDPureNDirectEarlyPrefix_20261127'),
    ("['CCoupledLoadContinuation', 'CCoupledSectionAnalysis']", "['CCoupledLoadContinuation']"),
    ('all254_other_components_unchanged=True', 'all255_other_components_unchanged=True'),
    ('new_direct_recovery_ASD_only=True', 'early_admissible_prefix_reconstructed_and_reused=True'),
    ('DIRECT ASD RECOVERY SOURCE VERIFIED;ONLY TWO EXPECTED COMPONENTS;MATERIALS AND TOLERANCES UNCHANGED',
     'DIRECT ASD EARLY PREFIX SOURCE VERIFIED;ONLY EXPECTED ONE-COMPONENT CHANGE')
]:
    assert code.count(before) == 1, before
    code = code.replace(before, after, 1)
start = code.index("patch = run_path(str(here / 'Patch-TestASDDirectRecovery.py'))")
end = code.index('for name in compiled[', start)
code = code[:start] + '''patch = run_path(str(here / 'Patch-TestDirectRecoveryPrefix.py'))
name = 'CCoupledLoadContinuation'
expected = patch['TestPatchDirectRecoveryPrefix'](body(old[name], name)[0], canonical=True)
assert normal(expected)[0] == normal(body(new[name], name)[0])[0]
''' + code[end:]
exec(compile(code, str(source), 'exec'), dict(__file__=str(Path(__file__).resolve()), __name__='TestCheckDirectEarlyPrefix'))
