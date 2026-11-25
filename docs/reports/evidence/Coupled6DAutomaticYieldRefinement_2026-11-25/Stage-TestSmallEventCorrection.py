"""ДЛЯ ТЕСТОВ. Читает остаточные поправки двух корней события782 на неизменной43-точечной истории."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import struct
import json

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold774ArcRootDerivative_20261125'
source_root = base / 'Cold782ProjectedEventRefinement_20261125'
root = base / 'Cold783SmallEventCorrection_20261125'
Reader = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']
r = Reader((source_root / 'Native/ArcYieldSplit.Result.bin').read_bytes()); v = [r.VectorTest() for _ in range(10)]
assert r.position == len(r.raw) and (v[0][:3] == 1).all()
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0, source_root / 'Input/Reverse1.bin')
source = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary'](source_root / 'Input/Reverse1.bin')
assert len(source['fields']) == 43
binary = bytearray()
for vector in (v[5],v[6],v[7],[v[0][3],v[0][4],v[0][10],v[0][11],0.,source['lambdas'][-1]]):
    binary += struct.pack('<i'+'d'*len(vector),len(vector),*vector)
values = {'Input/ArcRoots.bin':binary,
    'Input/SourceEvent782.bin':(source_root/'Native/ArcYieldSplit.Result.bin').read_bytes(),
    Path(__file__).name:Path(__file__).read_bytes()}
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in values.items():
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
prep['directed_scope']=dict(case='ActualSmallEventRootCorrection',source_states=43,roots_exact782=True,
    immutable_compiled774_reused=True,fresh_compilation_claimed=False,new_equilibrium_search=False,
    source_event_control_solver_not_called=True,working_residual_material_and_operator_methods_unchanged=True,
    physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('783 SMALL EVENT FULL/HALF ROOT CORRECTIONS; EXACT43 HISTORY;492 READONLY DIFFERENCES; COMPILED774 REUSED')
