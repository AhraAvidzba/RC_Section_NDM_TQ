"""ДЛЯ ТЕСТОВ. Повторяет чистое растяжение с имеющимся QR-резервом; изменяет только его доступность в отдельной книге."""
from pathlib import Path

here = Path(__file__).resolve().parent
source = here / 'Stage-TestASDHybridRecovery.py'
code = source.read_text('utf-8-sig')
assert code.count('Cold1030ASDScale1M_20261126') == 1
assert code.count('Cold1033ASDHybridRecovery_20261126') == 1
code = code.replace('Cold1030ASDScale1M_20261126', 'Cold1039ASDPureNTension_20261127')
code = code.replace('Cold1033ASDHybridRecovery_20261126', 'Cold1041ASDPureNHybrid_20261127')
code = code.replace('Run-TestASDBendingMainCapacity.ps1', 'Run-TestASDSpatialMainCapacity.ps1')
code = code.replace("base_Mx_Nmm=1000000.", "base_N_N=100000.,load_path='lambda*N'")
code = code.replace('PREPARED1033;', 'PREPARED1041;PURE N TENSION;')
exec(compile(code, str(source), 'exec'), dict(__file__=str(Path(__file__).resolve()), __name__='TestStageAxialHybrid'))
