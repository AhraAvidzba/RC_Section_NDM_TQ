"""ДЛЯ ТЕСТОВ. Запускает один исходный HardShear с нуля после восьми случаев и164 регрессий согласованного восстановления."""
from pathlib import Path
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
proof=json.loads((base/'Cold918AffineBranchRecovery_20261126/AffineBranchRecoveryProof.json').read_bytes())
assert not proof['lost_original_cases'] and proof['actual_prospective_diagnostic_records']>0
assert any(x['case']==7 and x['full_root_found'] for x in proof['cases'])
file=here/'Stage-TestStrictGrowthCold96.py';code=file.read_text('utf-8-sig')
for old,new in [('Cold888StrictEnvelopeGrowth_20261126','Cold918AffineBranchRecovery_20261126'),
 ('Cold890StrictGrowthCold96_20261126','Cold920AffineRecoveryCold96_20261126'),
 ('Cold889StrictGrowthRegressions_20261126','Cold919AffineRecoveryRegressions_20261126'),
 ("audit['imported_this_pass']==['CCoupledDeformationContinuation']","audit['imported_this_pass']==['CCoupledNewtonSolver']"),
 ('StrictGrowthTrueCold96','AffineRecoveryTrueCold96'),('exact_compiled_source888','exact_compiled_source918'),
 ('ACTUAL COLD;EXACT888;','ACTUAL COLD;EXACT918;')]:
    assert old in code,old;code=code.replace(old,new)
exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'affine_recovery_true_cold96_staging'})
