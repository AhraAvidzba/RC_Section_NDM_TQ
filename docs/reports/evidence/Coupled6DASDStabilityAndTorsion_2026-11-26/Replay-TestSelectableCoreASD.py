"""ДЛЯ ТЕСТОВ. Сверяет девять материальных шагов реально собранного рабочего ASD с нативным OpenSees."""
from pathlib import Path

here=Path(__file__).resolve().parent;source=here/'Replay-TestASDCurvatureMaterial.py';code=source.read_text('utf-8-sig')
code=code.replace('CurvatureMaterialReplayInput.json','SelectableCoreMaterialReplayInput.json')
code=code.replace('OpenSeesCurvatureMaterialReplay.json','OpenSeesSelectableCoreMaterialReplay.json')
code=code.replace('native_single_update_test','native_single_update_working')
code=code.replace('native_curve_sha256','native_binary_sha256').replace('PureBendingCurvature.bin','SelectableCore.bin')
exec(compile(code,str(source),'exec'),dict(__file__=str(Path(__file__).resolve()),__name__='working_ASD_replay_test'))
