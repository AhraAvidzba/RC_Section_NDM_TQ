"""ДЛЯ ТЕСТОВ. Сохраняет RPC-сбой компилятора до расчёта и повторяет подготовку в новом собственном Excel без изменения рабочего ядра."""
from pathlib import Path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');failed=base/'Cold1020ASDBendingMainCapacity_20261126'
failure=json.loads((failed/'CompilePointFailure.json').read_text('utf-8-sig'));assert '0x80010100' in failure['Message']
closure=json.loads((failed/'CompilePointClosure.json').read_text('utf-8-sig'));assert closure['State']=='Absent'
parent=json.loads((failed/'ParentCodeUpdate.json').read_bytes())
assert sha256((failed/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==parent['WorkbookSHA256']
proof=dict(preparation_failed=True,RPC_system_call_failed=True,own_compiler_closed=True,source_book_unchanged=True,
    numerical_calculation_not_started=True,not_a_material_or_equilibrium_failure=True,corrected_repeat='1021')
(failed/'CompilerRPCFailureProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
source=here/'Stage-TestASDBendingMainCapacity.py';code=source.read_text('utf-8-sig')
code=code.replace('Cold1020ASDBendingMainCapacity_20261126','Cold1021ASDBendingMainCapacityRetry_20261126').replace('PREPARED1020','PREPARED1021')
exec(compile(code,str(source),'exec'),dict(__file__=str(Path(__file__).resolve()),__name__='ASD_main_capacity_retry_stage_test'))
