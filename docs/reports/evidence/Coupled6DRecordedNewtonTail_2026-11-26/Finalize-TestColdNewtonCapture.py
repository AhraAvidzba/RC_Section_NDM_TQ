"""ДЛЯ ТЕСТОВ. Проверяет фактический проект257: новый автономный тестовый модуль, три перенаправленные записи и255 сохранённых компонентов."""
from pathlib import Path
from argparse import ArgumentParser
import json
repo=Path.cwd();here=Path(__file__).resolve().parent
p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,required=True);root=p.parse_args().root
imports=['CCoupledNewtonSolver','modTestColdNewtonCapture']
assert json.loads((root/'Preparation.json').read_bytes())['directed_scope']['actual_import_this_pass']==imports
source=here/'Finalize-TestStockYield.py';code=source.read_text('utf-8-sig')
old='assert set(old)==set(new) and len(new)==253';assert code.count(old)==1
code=code.replace(old,"assert len(old)==256 and set(new)==set(old)|{'modTestColdNewtonCapture'} and len(new)==257",1)
assert code.count('len(unchanged)==252')==1;code=code.replace('len(unchanged)==252','len(unchanged)==255',1)
old="compilation['Modules']==['CCoupledDeformationContinuation']";assert code.count(old)==2
code=code.replace(old,"compilation['Modules']=="+repr(imports))
code=code.replace("audit['imported_this_pass']=['CCoupledDeformationContinuation']","audit['imported_this_pass']=compilation['Modules']",1)
code=code.replace('actual_other_252_components_exact_except_identifier_case=True','actual_other_255_components_exact_except_identifier_case=True')
code=code.replace('STOCK YIELD CONTROLLER SOURCE VERIFIED; ONE CONTROLLER IMPORT, 252 PARENT COMPONENTS EXACT EXCEPT VBE IDENTIFIER CASE','COLD CAPTURE SOURCE VERIFIED; TWO IMPORTS, 255 OTHER PARENT COMPONENTS EXACT EXCEPT VBE IDENTIFIER CASE')
namespace={'__file__':str(source),'__name__':'cold_newton_capture_actual257_test'}
exec(compile(code,str(source),'exec'),namespace);namespace['TestFinalize'](repo,root)
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
assert audit['components']==257 and audit['imported_this_pass']==imports
assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==255
parent=Path(json.loads((root/'CompiledSourceReuse.json').read_bytes())['parent'])
name='src/Solver/CCoupledNewtonSolver.cls'
assert (root/name).read_text('utf-8-sig').replace('modTestColdNewtonCapture.TestCapture','modTestCoupledElasticState.TestNewtonRejectedCapture')==(parent/name).read_text('utf-8-sig')
assert (root/'tests/modTestCoupledElasticState.bas').read_bytes()==(parent/'tests/modTestCoupledElasticState.bas').read_bytes()
print('ACTUAL257 VERIFIED;ONE NEW TEST MODULE;255 OTHER COMPONENTS PRESERVED;NUMERICAL BODY EXACT918')
