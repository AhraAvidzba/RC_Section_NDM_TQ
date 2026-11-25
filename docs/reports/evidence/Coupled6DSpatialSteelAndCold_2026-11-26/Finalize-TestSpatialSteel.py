"""ДЛЯ ТЕСТОВ. Проверяет шесть импортов с фактическим порядком каталогов Materials/Solver/tests и точностью всех 256 компонентов."""
from pathlib import Path
import json
here=Path(__file__).resolve().parent
root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold881SpatialSteelLimits_20261126')
expected=json.loads((root/'Preparation.json').read_bytes())['directed_scope']['actual_import_this_pass']
actual=json.loads((root/'docs/regression/CodeUpdate.json').read_bytes())['Modules']
assert len(actual)==len(set(actual))==6 and set(actual)==set(expected)
file=here/'Finalize-TestOperationalArc.py';code=file.read_text('utf-8-sig')
anchor='imports = sorted(imports)';assert code.count(anchor)==1
code=code.replace(anchor,"assert set(imports)==set(json.loads((root/'docs/regression/CodeUpdate.json').read_bytes())['Modules'])\nimports = json.loads((root/'docs/regression/CodeUpdate.json').read_bytes())['Modules']",1)
code=code.replace("default=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold785OperationalArcController_20261125')","default=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold881SpatialSteelLimits_20261126')")
exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'spatial_steel_source_test'})
