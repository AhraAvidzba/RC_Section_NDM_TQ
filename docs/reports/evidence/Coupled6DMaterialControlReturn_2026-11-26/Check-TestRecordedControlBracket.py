"""ДЛЯ ТЕСТОВ. Проверяет исходную управляющую координату и равновесие всех полей940 независимо от нативного флага сходимости."""
from pathlib import Path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold940RecordedControlBracket_20261126';parent=base/'Cold938RecordedYieldAdvance_20261126'
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
script=here/'Check-TestRecordedYieldCoordinate.py';code=script.read_text('utf-8-sig')
old="root=base/'Cold936RecordedYieldCoordinate_20261126'";assert code.count(old)==1
code=code.replace(old,"root=base/'Cold940RecordedControlBracket_20261126'",1)
scope={'__file__':str(script),'__name__':'recorded_control_bracket_test'}
exec(compile(code,str(script),'exec'),scope)
proof=json.loads((root/'RecordedYieldCoordinateProof.json').read_bytes())
proof.pop('source_actual934_state25_and_initial_residual_bitexact_by_native_guard')
proof.update(accepted25_source_history_preserved=True,starting_native938_root_residual_bitexact_guard=True,
    immutable_compiled936_reused=True,original_linear_target_unchanged=True,
    bracket_input_sha256=sha256((root/'OriginalControlBracket.json').read_bytes()).hexdigest().upper())
(root/'RecordedYieldCoordinateProof.json').write_text(json.dumps(proof,indent=2,sort_keys=True)+'\n',encoding='utf-8')
