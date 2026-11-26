"""ДЛЯ ТЕСТОВ. Подготавливает читающий опыт согласованности прогнозной касательной с её собственным направлением."""
from pathlib import Path
import json
here=Path(__file__).resolve().parent;file=here/'Stage-TestCold32DirectionalJacobian.py';code=file.read_text('utf-8-sig')
assert json.loads(Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold906Cold32OutgoingStressScan_20261126/Cold32OutgoingStressScanProof.json').read_bytes())['zero_increment_stress_drift_MPa']==0
begin=code.index("entry='''");end=code.index("\nname='tests/",begin)
code=code[:begin]+'entry='+repr((here/'TestCold32TangentConeIterations.bas.txt').read_text('utf-8-sig'))+code[end:]
code=code.replace('Cold899Cold32DirectionalJacobian_20261126','Cold907Cold32TangentConeIterations_20261126').replace('TestReadCold32DirectionalJacobian','TestReadCold32TangentConeIterations')
exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'cold32_tangent_cone_staging'})
root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold907Cold32TangentConeIterations_20261126')
prep=json.loads((root/'Preparation.json').read_bytes());scope=prep['directed_scope']
scope.pop('levels');scope.pop('signed_directional_steps')
scope.update(case='Cold32ReadOnlyTangentConeIterations',probe_denominators=[1024,64,8],iterations_per_probe=4,
    actual_directional_residual_steps=[0,1/1024,1/4096],all_probe_fields_diagnostic_not_accepted_history=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
