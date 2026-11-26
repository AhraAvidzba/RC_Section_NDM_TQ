"""ДЛЯ ТЕСТОВ. Использует подготовку предыдущей книги для проверки точной силовой пробы сначала от левого состояния ASD."""
from pathlib import Path

here = Path(__file__).resolve().parent
source = here / 'Stage-TestDirectRecoveryObservation.py'
code = source.read_text('utf-8-sig')
for before, after in [
    ("parent = base / 'Cold1062ASDPureNDirectRecovery_20261127'", "parent = base / 'Cold1063ASDPureNDirectObservation_20261127'"),
    ("root = base / 'Cold1063ASDPureNDirectObservation_20261127'", "root = base / 'Cold1064ASDPureNDirectWarmTrial_20261127'"),
    ('Patch-TestDirectRecoveryObservation.py', 'Patch-TestDirectWarmTrial.py'),
    ('TestPatchDirectRecoveryObservation', 'TestPatchDirectWarmTrial'),
    ('Stage-TestDirectRecoveryObservation.py', 'Stage-TestDirectWarmTrial.py'),
    ("variant='observe-direct'", "variant='direct-warm-trial'"),
    ("case='ASDDirectPureNObservation'", "case='ASDDirectPureNWarmTrial'"),
    ('only_reading_observations_added=True', 'unseeded_exact_trial_before_interpolated_trial=True'),
    ('numerical_algorithm_unchanged=True', 'original_materials_and_tolerances_unchanged=True'),
    ('DIRECT ASD OBSERVATION PREPARED;NUMERICAL EXPRESSIONS UNCHANGED', 'DIRECT ASD WARM EXACT TRIAL PREPARED;MATERIALS AND TOLERANCES UNCHANGED')
]:
    assert before in code, before
    code = code.replace(before, after)
start = code.index("compiler = (parent / 'Compile-TestASDDirectRecovery.ps1')")
end = code.index('for name in (', start)
code = code[:start] + '''shutil.copyfile(parent / 'Compile-TestDirectRecoveryObservation.ps1', root / 'Compile-TestDirectWarmTrial.ps1')
''' + code[end:]
exec(compile(code, str(source), 'exec'), dict(__file__=str(Path(__file__).resolve()), __name__='TestStageDirectWarmTrial'))
