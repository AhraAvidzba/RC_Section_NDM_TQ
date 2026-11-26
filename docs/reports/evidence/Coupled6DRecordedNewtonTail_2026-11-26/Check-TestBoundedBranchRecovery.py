"""ДЛЯ ТЕСТОВ. Проверяет точную систему54/55 и исходные критерии корней ограниченного временного корректора."""
from pathlib import Path
import json
here=Path(__file__).resolve().parent;file=here/'Check-TestRecordedTailSystem.py'
code=file.read_text('utf-8-sig')
for old,new in [
 ('Cold925RecordedTailSystem_20261126','Cold927BoundedBranchRecovery_20261126'),
 ('Cold922ColdNewtonCapture_20261126','Cold925RecordedTailSystem_20261126'),
 ("audit['imported_this_pass']==['modTestCoupledElasticState']","audit['imported_this_pass']==['CCoupledNewtonSolver']"),
 ("if file.is_file():assert file.read_bytes()==(root/file.relative_to(parent)).read_bytes()",
  "if file.is_file() and file.name!='CCoupledNewtonSolver.cls':assert file.read_bytes()==(root/file.relative_to(parent)).read_bytes()"),
 ('RecordedTailSystemProof.json','BoundedBranchRecoveryProof.json')]:
    assert old in code,old;code=code.replace(old,new)
scope={'__file__':str(file),'__name__':'bounded_branch_recovery_test'}
exec(compile(code,str(file),'exec'),scope)
root=scope['root'];parent=scope['parent'];proof=scope['proof']
assert (root/'tests/modTestCoupledElasticState.bas').read_bytes()==(parent/'tests/modTestCoupledElasticState.bas').read_bytes()
code=(root/'src/Solver/CCoupledNewtonSolver.cls').read_text('utf-8-sig')
assert 'For localStep = 1 To 3' in code and 'If mIterations + 1 >= maxIterations Then GoTo RestoreFailure' in code
assert 'mIterations = mIterations + 1' in code[code.index('Private Function TryProspectiveBranchStep('):]
proof.update(temporary_branch_steps_limited_to_three=True,
    all_additional_steps_inside_original30_budget=True,original_material_and_force_acceptance_unchanged=True,
    all256_other_source_components_preserved=True,diagnostic_last_norm_is_restored_context_not_minimum_over_all_attempts=True,
    no_working_or_release_promotion=True)
(root/'BoundedBranchRecoveryProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
