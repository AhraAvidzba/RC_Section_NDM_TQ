"""ДЛЯ ТЕСТОВ. Проверяет согласованный шаг активной ветви по тем же восьми случаям и независимой сборке корней."""
from pathlib import Path
here=Path(__file__).resolve().parent;file=here/'Check-TestProspectiveBranchRecovery.py';code=file.read_text('utf-8-sig')
for old,new in [('Cold915ProspectiveBranchRecovery_20261126','Cold918AffineBranchRecovery_20261126'),
 ('Cold911NewtonRejectedCapture_20261126','Cold915ProspectiveBranchRecovery_20261126'),
 ('Stage-TestProspectiveBranchRecovery.py','Stage-TestAffineBranchRecovery.py'),
 ('ProspectiveBranchRecoveryProof.json','AffineBranchRecoveryProof.json'),
 ('only_prospective_branch_recovery_added=True','only_matched_affine_branch_recovery_added=True')]:
    assert old in code,old;code=code.replace(old,new)
exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'affine_branch_recovery_roots_test'})
