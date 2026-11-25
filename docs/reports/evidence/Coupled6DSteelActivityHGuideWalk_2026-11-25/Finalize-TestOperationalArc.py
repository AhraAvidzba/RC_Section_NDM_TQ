"""ДЛЯ ТЕСТОВ. Проверяет явный набор до шести импортов и точность остальных компонентов фактического проекта из 256 компонентов."""
from pathlib import Path
from runpy import run_path
import json
from argparse import ArgumentParser

repo = Path.cwd(); here = Path(__file__).resolve().parent
p = ArgumentParser(description=__doc__)
p.add_argument('--root', type=Path, default=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold785OperationalArcController_20261125'))
root = p.parse_args().root
imports = json.loads((root / 'Preparation.json').read_bytes())['directed_scope']['actual_import_this_pass']
assert len(imports) == len(set(imports)) and 1 <= len(imports) <= 6
imports = sorted(imports)
source = here / 'Finalize-TestStockYield.py'
code = source.read_text('utf-8-sig')
assert code.count('assert set(old)==set(new) and len(new)==253') == 1
code = code.replace('assert set(old)==set(new) and len(new)==253', "assert set(new)==set(old)|{'CCoupledArcContinuation'} and len(new)==256")
code = code.replace('len(unchanged)==252', 'len(unchanged)=='+str(256-len(imports)))
before = "compilation['Modules']==['CCoupledDeformationContinuation']"; assert code.count(before) == 2
code = code.replace(before, "compilation['Modules']==" + repr(imports))
code = code.replace("audit['imported_this_pass']=['CCoupledDeformationContinuation']", "audit['imported_this_pass']=compilation['Modules']", 1)
namespace = {'__file__':str(source), '__name__':'operational_arc_actual256_test'}
exec(compile(code, str(source), 'exec'), namespace)
namespace['TestFinalize'](repo, root)
audit = json.loads((root / 'PointTraceSourceComposition.json').read_bytes())
assert audit['components'] == 256 and audit['imported_this_pass'] == imports
assert len(audit['other_actual_components_exact_parent_except_identifier_case']) == 256-len(imports)
print('ACTUAL256 VERIFIED; IMPORTS',len(imports),'OTHER PARENT COMPONENTS VERIFIED',256-len(imports))
