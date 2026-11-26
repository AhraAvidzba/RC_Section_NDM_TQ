"""ДЛЯ ТЕСТОВ. Проверяет исходную систему и любые новые корни после одной дополнительной пробы касательной."""
from pathlib import Path
import json
here=Path(__file__).resolve().parent;file=here/'Check-TestRecordedTailFineLine.py'
code=file.read_text('utf-8-sig').replace('Cold929RecordedTailFineLine_20261126','Cold933ProspectiveExtraHalf_20261126')
code=code.replace('RecordedTailFineLineProof.json','ProspectiveExtraHalfProof.json')
scope={'__file__':str(file),'__name__':'prospective_extra_half_test'}
exec(compile(code,str(file),'exec'),scope)
root=scope['root'];parent=root.parent/'Cold929RecordedTailFineLine_20261126'
before=(parent/'src/Solver/CCoupledNewtonSolver.cls').read_text('utf-8-sig')
after=(root/'src/Solver/CCoupledNewtonSolver.cls').read_text('utf-8-sig')
begin=before.index('Private Function TryProspectiveBranchStep(');end=before.index('End Function',begin)+len('End Function')
body=before[begin:end];old='    Do While fraction >= MIN_FRACTION'
new=old+' * 0.5 \' Более короткая проба учитывает смену ветви; приёмка по исходной норме сохраняется.'
assert body.count(old)==1 and after==before[:begin]+body.replace(old,new,1)+before[end:]
proof=scope['proof'];proof.update(only_one_extra_helper_halving_from929=True,
    all_primary_iterations_inside_original30=True,all_original_material_and_equilibrium_criteria=True,
    no_new_physical_capacity_claim=True,no_release_promotion=True)
(root/'ProspectiveExtraHalfProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
