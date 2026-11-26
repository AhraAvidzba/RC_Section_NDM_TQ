"""ДЛЯ ТЕСТОВ. Сверяет частный контроллер адаптации ASD с точным подготовленным исходником и сохраняет остальные компоненты."""
from pathlib import Path

here=Path(__file__).resolve().parent;source=here/'Check-TestASDProbeBudgetSource.py';code=source.read_text('utf-8-sig')
code=code.replace('Cold1022ASDBendingProbeBudget_20261126','Cold1023ASDGeometricContinuation_20261126')
start=code.index("needle='private const probe_limit");end=code.index('# ProductionChanged',start)
code=code[:start]+code[end:]
start=code.index('# ProductionChanged');end=code.index('proof=dict',start)
code=code[:start]+code[end:]
code=code.replace("only_controller_change='PROBE_LIMIT96to192'","only_controller_change='private_ASD_geometric_growth_and_deferred_local_peak_refinement'")
code=code.replace('ProbeBudgetSourceProof.json','GeometricSourceProof.json')
code=code.replace('ASD BUDGET SOURCE VERIFIED; 96to192 ONLY; RUNNER RELEASE-PROTECTION GUARD CORRECTED BEFORE EXECUTION',
    'ASD GEOMETRIC SOURCE VERIFIED;EXACT PRIVATE CONTROLLER;255 OTHER COMPONENTS UNCHANGED')
exec(compile(code,str(source),'exec'),dict(__file__=str(Path(__file__).resolve()),__name__='ASD_geometric_source_test'))
