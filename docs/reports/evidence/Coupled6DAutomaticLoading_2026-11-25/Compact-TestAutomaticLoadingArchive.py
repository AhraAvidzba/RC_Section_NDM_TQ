"""ДЛЯ ТЕСТОВ. Убирает из нового архива только побитно подтверждённые копии неиспользованных унаследованных входов; исходники на D сохраняются."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json,ast

repo=Path.cwd().resolve();here=Path(__file__).resolve().parent
out=(repo/'docs/reports/evidence/Coupled6DAutomaticLoading_2026-11-25').resolve()
assert out.is_relative_to(repo) and out.name=='Coupled6DAutomaticLoading_2026-11-25'
tree=ast.parse((here/'Record-TestAutomaticLoading.py').read_text('utf-8-sig'))
names=next(ast.literal_eval(n.value) for n in tree.body if isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id=='names' for t in n.targets))
stage_helpers=next(ast.literal_eval(n.value) for n in tree.body if isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id=='stage_helpers' for t in n.targets))
keep={'Reverse1.bin','SourceColdRawHistory.bin','ArcCallAllowance.bin','OperationalArcRatio.bin','NativeTolerance.bin','NativeSolver.json',
    'EnvelopeOptions.bin','HPathTestOptions.bin','ColdHPrefix.bin','ColdHRequest.bin','MaterialProbeRatio.bin'}
old=out/'OmittedInheritedInputs.json'
omitted=json.loads(old.read_bytes()) if old.exists() else {}
for k,name in names.items():
    source=Path('D:/CodexScratch/RC_Section_NDM_TQ')/name/'Input'
    for p in (out/str(k)/'Input').iterdir():
        assert p.is_file() and p.resolve().is_relative_to(out)
        if p.name not in keep:
            data=p.read_bytes();assert data==(source/p.name).read_bytes()
            omitted[p.relative_to(out).as_posix()]=dict(sha256=sha256(data).hexdigest().upper(),bytes=len(data),preserved_source=str(source/p.name))
            p.unlink()
    for p in (out/str(k)).glob('Stage-Test*.py'):
        assert p.is_file() and p.resolve().is_relative_to(out)
        if p.name!=stage_helpers[k]:
            origin=source.parent/p.name;data=p.read_bytes();assert data==origin.read_bytes()
            omitted[p.relative_to(out).as_posix()]=dict(sha256=sha256(data).hexdigest().upper(),bytes=len(data),preserved_source=str(origin))
            p.unlink()
(out/'OmittedInheritedCopies.json').write_text(json.dumps(omitted,indent=2,sort_keys=True)+'\n',encoding='utf-8')
if old.exists():old.unlink()
for n in ('Record-TestAutomaticLoading.py',Path(__file__).name):(out/n).write_bytes((here/n).read_bytes())
pins={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file() and p.name!='Files.json'}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('COPIES OF UNUSED INHERITED INPUTS REMOVED',len(omitted),'ALL SOURCES PRESERVED ON D; ARCHIVE PINS',len(pins))
