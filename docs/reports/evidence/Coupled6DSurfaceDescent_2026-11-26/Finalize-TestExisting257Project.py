"""ДЛЯ ТЕСТОВ. Проверяет объявленные импорты в уже существующем257-компонентном проекте без добавления компонентов."""
from pathlib import Path
from argparse import ArgumentParser
import json
repo=Path.cwd();here=Path(__file__).resolve().parent;p=ArgumentParser(description=__doc__)
p.add_argument('--root',type=Path,required=True);root=p.parse_args().root
imports=sorted(json.loads((root/'Preparation.json').read_bytes())['directed_scope']['actual_import_this_pass'])
assert 1<=len(imports)<=6 and len(imports)==len(set(imports))
source=here/'Finalize-TestStockYield.py';code=source.read_text('utf-8-sig')
old='assert set(old)==set(new) and len(new)==253';assert code.count(old)==1
code=code.replace(old,'assert set(old)==set(new) and len(new)==257',1)
code=code.replace('len(unchanged)==252','len(unchanged)=='+str(257-len(imports)),1)
old="compilation['Modules']==['CCoupledDeformationContinuation']";assert code.count(old)==2
code=code.replace(old,"compilation['Modules']=="+repr(imports))
code=code.replace("audit['imported_this_pass']=['CCoupledDeformationContinuation']","audit['imported_this_pass']=compilation['Modules']",1)
code=code.replace('actual_other_252_components_exact_except_identifier_case=True','actual_other_'+str(257-len(imports))+'_components_exact_except_identifier_case=True')
code=code.replace('STOCK YIELD CONTROLLER SOURCE VERIFIED; ONE CONTROLLER IMPORT, 252 PARENT COMPONENTS EXACT EXCEPT VBE IDENTIFIER CASE','EXISTING257 SOURCE VERIFIED; DECLARED IMPORTS AND OTHER PARENT COMPONENTS CHECKED')
namespace={'__file__':str(source),'__name__':'existing257_declared_imports_test'}
exec(compile(code,str(source),'exec'),namespace);namespace['TestFinalize'](repo,root)
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
assert audit['components']==257 and audit['imported_this_pass']==imports
assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==257-len(imports)
print('EXISTING257 VERIFIED;DECLARED IMPORTS',imports,'OTHER PARENT COMPONENTS',257-len(imports))
