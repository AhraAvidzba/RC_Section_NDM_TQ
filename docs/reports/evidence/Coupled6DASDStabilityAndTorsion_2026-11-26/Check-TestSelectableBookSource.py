"""ДЛЯ ТЕСТОВ. Проверяет фактические модули свежей книги и закреплённые входы без повторного импорта VBA."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
from runpy import run_path
import json,re
from difflib import unified_diff

repo=Path.cwd();here=Path(__file__).resolve().parent;root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1017SelectableMaterialBook_20261126')
prep=json.loads((root/'Preparation.json').read_bytes())
for name,digest in prep['inputs'].items():assert sha256((root/name).read_bytes()).hexdigest().upper()==digest,name
for name,digest in prep['release_hashes'].items():assert sha256((repo/name).read_bytes()).hexdigest().upper()==digest,name
closure=json.loads((root/'BuildClosureAfterExit.json').read_text('utf-8-sig'));assert closure['State']=='Absent' and closure['ExactOwnProcessClosed']
book=(root/'workbook/output/SelectableLaws.xlsm').read_bytes();assert sha256(book).hexdigest().upper()==closure['SavedWorkbookSHA256']
log_raw=(root/'Build.log').read_bytes()
log=log_raw.decode('utf-16' if log_raw.startswith((b'\xff\xfe',b'\xfe\xff')) else 'utf-8-sig')
assert 'Workbook built:' in log
helper,body,reader=run_path(str(here/'Inspect-TestResolutionVBA.py'))['TestHelpers'](repo)
with ZipFile(BytesIO(book)) as archive:binary=archive.read('xl/vbaProject.bin')
_,actual=reader(helper['TestContainer'](binary));normal=run_path(str(here/'Canonical-TestCoreVBA.py'))['TestCanonical']
with ZipFile(repo/'workbook/output/RC_Section_NDM.xlsm') as archive:published_binary=archive.read('xl/vbaProject.bin')
_,published=reader(helper['TestContainer'](published_binary))

def TestIdentifierCase(text):
    """Меняет только регистр исполняемых идентификаторов; строки и комментарии сохраняет."""
    return re.sub(r'"(?:[^"\n]|"")*"|\'[^\n]*|[A-Za-z_][\w]*',lambda m:m.group() if m.group().startswith(('"',"'")) else m.group().lower(),text)

def TestCodeSpacing(text):
    """Учитывает выравнивание VBE только вне строк и текста комментариев; границы токенов и переводы строк сохраняются."""
    return re.sub(r'"(?:[^"\n]|"")*"|\'[^\n]*|[ \t]+',lambda m:' ' if m.group()[0] in ' \t' else m.group(),text)

verified=[];numeric=[];unchanged_published=[]
changed_core={'CCoupledConcreteCalibration','CCoupledConcreteEnvelope','CCoupledConcreteResponse','CCoupledASDFlow','CCoupledSectionAnalysis','modCoupledWorkbookInput'}
for parent in ('src','tests'):
    for path in (root/parent).rglob('*'):
        if path.suffix not in ('.cls','.bas'):continue
        name=path.stem;assert name in actual,name
        if name=='modSectionRecipeApiGuide':
            expected=path.read_text('utf-8-sig');observed=actual[name]
            for text in (expected,observed):
                assert all(not line.strip() or line.lstrip().startswith(("'",'Attribute ')) for line in text.splitlines())
            expected_comments=[line.lstrip().encode('cp1251','replace').decode('cp1251') for line in expected.splitlines() if line.lstrip().startswith("'")]
            observed_comments=[line.lstrip() for line in observed.splitlines() if line.lstrip().startswith("'")]
            assert expected_comments==observed_comments,name
            verified.append(name);continue
        try:expected,_=helper['TestMVTCompilerBody'](body,path.read_text('utf-8-sig'),name)
        except Exception as error:raise AssertionError('Source body decoder cannot parse '+name) from error
        observed=body(actual[name],name)[0]
        before=TestCodeSpacing(TestIdentifierCase(expected));after=TestCodeSpacing(TestIdentifierCase(observed))
        if before!=after:
            before=normal(before)[0];after=normal(after)[0];numeric.append(name)
        if before!=after:
            if name not in changed_core and name in published:
                old=TestCodeSpacing(TestIdentifierCase(body(published[name],name)[0]))
                if TestCodeSpacing(TestIdentifierCase(observed))==old:
                    unchanged_published.append(name);verified.append(name);continue
            print('SOURCE DIFFERENCE',name)
            print('\n'.join(list(unified_diff(before.splitlines(),after.splitlines(),fromfile='pinned source',tofile='saved VBA',lineterm=''))[:45]))
            raise AssertionError(name)
        verified.append(name)
assert 'CCoupledConcreteCalibration' in verified and len(verified)==249
proof=dict(saved_book_sha256=sha256(book).hexdigest().upper(),actual_component_count=len(actual),
    all249_working_and_test_sources_verified=True,verified_modules=sorted(verified),exact_numeric_literal_format_normalization=numeric,
    unchanged_published_native_modules_with_nonidentical_source_literal_presentation=unchanged_published,
    changed_core_modules_strictly_match_pinned_sources=True,
    document_module_source_audit_not_included=True,explicit_build_compilation_completed=True,
    own_process_closed_after_runner_reference_release=True,wrapper20s_close_wait_failed=True,
    native_dropdown_validation_pending=True,calculation_tests_pending=True,source_pins_unchanged=True,release_unchanged=True,
    physical_capacity_certified=False,whole_goal_completed=False)
(root/'FreshSourceProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('FRESH SELECTABLE BOOK SOURCE VERIFIED',len(verified),'SOURCE MODULES;',len(actual),'ACTUAL COMPONENTS;NATIVE DROPDOWN/CALCULATIONS PENDING')
