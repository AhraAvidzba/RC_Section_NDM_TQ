"""ДЛЯ ТЕСТОВ. Применяет неизменённую независимую приёмку958 к усиленному внутреннему решению959."""
from pathlib import Path
import json
here=Path(__file__).resolve().parent;script=here/'Check-TestBoundedIncomingSteps.py';code=script.read_text('utf-8-sig')
changes={
 "root=base/'Cold958BoundedIncomingSteps_20261126';parent=base/'Cold957FourPartEventHistory_20261126'":"root=base/'Cold959AccurateIncomingSteps_20261126';parent=base/'Cold958BoundedIncomingSteps_20261126'",
 "audit['imported_this_pass']==['modTestCoupledElasticState']":"audit['imported_this_pass']==['CCoupledNewtonSolver']",
 "if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()":"if path.is_file() and path.name!='CCoupledNewtonSolver.cls':assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()",
}
for old,new in changes.items():assert code.count(old)==1;code=code.replace(old,new,1)
env={'__file__':str(script),'__name__':'accurate_incoming_steps_test'};exec(compile(code,str(script),'exec'),env)
root=env['root'];parent=env['parent']
assert (root/'tests/modTestCoupledElasticState.bas').read_bytes()==(parent/'tests/modTestCoupledElasticState.bas').read_bytes()
assert (root/'Input/BoundedIncoming.bin').read_bytes()==(parent/'Input/BoundedIncoming.bin').read_bytes()
before=(parent/'src/Solver/CCoupledNewtonSolver.cls').read_text('utf-8-sig')
after=(root/'src/Solver/CCoupledNewtonSolver.cls').read_text('utf-8-sig')
line="Private Const MATERIAL_RETURN_SOLVE_FRACTION As Double = 0.0001 ' Уточняет только внутренние равновесия возврата; приёмка формы, сил и оценка неопределённости используют исходную точность.\n"
assert after.count(line)==1
restored=after.replace(line,'',1)
old='absoluteTolerance * MATERIAL_RETURN_SOLVE_FRACTION, relativeTolerance * MATERIAL_RETURN_SOLVE_FRACTION, remaining, consistentControl:='
assert restored.count(old)==2
restored=restored.replace(old,'absoluteTolerance, relativeTolerance, remaining, consistentControl:=')
assert restored==before
proof=json.loads((root/'BoundedIncomingStepsProof.json').read_bytes())
proof.update(all_numerical_sources_exact957=False,all256_other_components_preserved=True,only_two_internal_solve_tolerance_calls_changed=True,
    internal_solve_tolerance_factor=.0001,original_force_half_lambda_noise_uncertainty_and_steel_guards_bitexact958=True,
    test_entry_and_input_bitexact958=True)
for name in ('BoundedIncomingStepsProof.json','AccurateIncomingStepsProof.json'):
    (root/name).write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
