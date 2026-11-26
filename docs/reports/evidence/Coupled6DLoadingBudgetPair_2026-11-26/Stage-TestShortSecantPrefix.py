"""ДЛЯ ТЕСТОВ. Переиспользует точную собранную965, начиная с доли1/64 без повторения больших неудачных префиксов."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold965ControlledSecantPrefix_20261126';root=base/'Cold966ShortSecantPrefix_20261126'
assert not json.loads((parent/'ControlledSecantPrefixProof.json').read_bytes())['returned_verified_candidate']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((parent/'Input/ControlledPrefix.bin').read_bytes());r.position=0
context,p=r.VectorTest(),r.VectorTest();assert r.position==len(r.raw);context[-1]=1/64
vec=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector'];name='Input/ControlledPrefix.bin'
(root/name).write_bytes(vec(context)+vec(p));prep=json.loads((root/'Preparation.json').read_bytes())
prep['inputs'][name]=sha256((root/name).read_bytes()).hexdigest().upper()
(root/Path(__file__).name).write_bytes(Path(__file__).read_bytes());prep['inputs'][Path(__file__).name]=sha256(Path(__file__).read_bytes()).hexdigest().upper()
prep['directed_scope'].update(case='SameCompiledSecantPrefixesStartingAt1Over64',actual_import_this_pass=[],all257_actual_components_and_saved_book_bitexact965=True,
    initial_fraction=1/64,no_new_numerical_code_or_compilation=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'ALL257 SOURCES AND SAVED BOOK EXACT965;INITIAL FRACTION1/64;SHARED30 UNCHANGED')
