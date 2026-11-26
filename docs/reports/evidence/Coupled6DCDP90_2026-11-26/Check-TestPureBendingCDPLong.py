"""ДЛЯ ТЕСТОВ. Проверяет завершённые состояния CDP с явно увеличенным сторожевым лимитом 1800 с, не меняя силовые критерии."""
from pathlib import Path
import ast,json,re,sys

here=Path(__file__).resolve().parent;source=here/'Check-TestPersistentHEvidence.py'
tree=ast.parse(source.read_text('utf-8-sig'));method=next(n for n in tree.body if isinstance(n,ast.FunctionDef) and n.name=='TestNative')
native_code=ast.get_source_segment(source.read_text('utf-8-sig'),method)
assert native_code.count('timeout=600 sec;')==1 and native_code.count('timeout600')==1
native_code=native_code.replace('timeout=600 sec;','timeout=1800 sec;').replace('timeout600','timeout1800')
namespace=dict(json=json,re=re);exec(compile(native_code,str(source),'exec'),namespace)
checker=here/'Check-TestPureBendingCurvature.py';code=checker.read_text('utf-8-sig')
call="native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')"
assert code.count(call)==1
code=code.replace(call,"native,text=TestLongNative(root,'RealColdReplay')",1).replace('timeout600','timeout1800')
root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1014PureBendingCDPLong_20261126')
prep=json.loads((root/'Preparation.json').read_bytes());assert prep['directed_scope']['watchdog_seconds']==1800
sys.argv=[str(checker),'--root',str(root),'--allow-timeout-prefix']
exec(compile(code,str(checker),'exec'),dict(__file__=str(checker),__name__='long_CDP_curve_check_test',TestLongNative=namespace['TestNative']))
