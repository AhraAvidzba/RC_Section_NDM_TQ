"""ДЛЯ ТЕСТОВ. Подготавливает прежние шесть холодных осевых случаев ASD для восстановления от раннего допустимого префикса."""
from pathlib import Path

here = Path(__file__).resolve().parent
source = here / 'Stage-TestDirectRecoveryObservation.py'
code = source.read_text('utf-8-sig')
for before, after in [
    ("parent = base / 'Cold1062ASDPureNDirectRecovery_20261127'", "parent = base / 'Cold1064ASDPureNDirectWarmTrial_20261127'"),
    ("root = base / 'Cold1063ASDPureNDirectObservation_20261127'", "root = base / 'Cold1065ASDPureNDirectEarlyPrefix_20261127'"),
    ('Patch-TestDirectRecoveryObservation.py', 'Patch-TestDirectRecoveryPrefix.py'),
    ('TestPatchDirectRecoveryObservation', 'TestPatchDirectRecoveryPrefix'),
    ('Stage-TestDirectRecoveryObservation.py', 'Stage-TestDirectRecoveryPrefix.py'),
    ("variant='observe-direct'", "variant='direct-early-prefix'"),
    ("case='ASDDirectPureNObservation'", "case='ASDDirectPureNEarlyPrefix'"),
    ('only_reading_observations_added=True', 'early_admissible_prefix_reconstructed_and_reused=True'),
    ('numerical_algorithm_unchanged=True', 'original_materials_and_tolerances_unchanged=True'),
    ('DIRECT ASD OBSERVATION PREPARED;NUMERICAL EXPRESSIONS UNCHANGED', 'DIRECT ASD EARLY PREFIX PREPARED;EXACT SEED POINT RESOLVED;MATERIALS AND TOLERANCES UNCHANGED')
]:
    assert before in code, before
    code = code.replace(before, after)
start = code.index("compiler = (parent / 'Compile-TestASDDirectRecovery.ps1')")
end = code.index('for name in (', start)
code = code[:start] + '''shutil.copyfile(parent / 'Compile-TestDirectWarmTrial.ps1', root / 'Compile-TestDirectRecoveryPrefix.ps1')
''' + code[end:]
exec(compile(code, str(source), 'exec'), dict(__file__=str(Path(__file__).resolve()), __name__='TestStageDirectEarlyPrefix'))
