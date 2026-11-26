"""ДЛЯ ТЕСТОВ. Проверяет однокомпонентную сборку относительно точного родительского VBA-проекта."""

from argparse import ArgumentParser
from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path
import re
from runpy import run_path
from zipfile import ZipFile


def TestFinalize(repo,root):
    """Проверяет 252 остальных компонента с явным учётом регистра имён VBE; строки, комментарии и остальные символы точны."""
    here=Path(__file__).resolve().parent
    compilation=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
    assert compilation['Compiled'] and not compilation['ProductionChanged']
    assert compilation['Modules']==['CCoupledDeformationContinuation']
    closure=(root/'CompilePointClosure.json').read_bytes()
    assert json.loads(closure)['State']=='Absent'
    parent=json.loads((root/'ParentCodeUpdate.json').read_text('utf-8-sig'))
    source_root=Path(json.loads((root/'CompiledSourceReuse.json').read_bytes())['parent'])
    source=(source_root/'docs/regression/Diagnostic.xlsm').read_bytes()
    assert sha256(source).hexdigest().upper()==parent['WorkbookSHA256']
    actual=(root/'docs/regression/Diagnostic.xlsm').read_bytes()
    assert sha256(actual).hexdigest().upper()==compilation['WorkbookSHA256']
    with ZipFile(BytesIO(source)) as book:before=book.read('xl/vbaProject.bin')
    with ZipFile(BytesIO(actual)) as book:after=book.read('xl/vbaProject.bin')
    helper,_,reader=run_path(str(here/'Inspect-TestResolutionVBA.py'))['TestHelpers'](repo)
    _,old=reader(helper['TestContainer'](before));_,new=reader(helper['TestContainer'](after))
    assert set(old)==set(new) and len(new)==253
    unchanged=[name for name in new if name not in compilation['Modules']]
    different=[name for name in unchanged if old[name]!=new[name]]
    def TestIdentifierCase(text):
        """Учитывает только регистр исполняемых имён, не меняя строки и комментарии."""
        return re.sub(r'"(?:[^"\n]|"")*"|\'[^\n]*|[A-Za-z_][\w]*',
            lambda item:item.group() if item.group().startswith(('"',"'")) else item.group().lower(),text)
    assert len(unchanged)==252
    assert all(TestIdentifierCase(old[name])==TestIdentifierCase(new[name]) for name in different),different
    source_file=here/'Audit-TestPrivateCoupledBook.py'
    code=source_file.read_text('utf-8-sig')
    anchor="    assert compilation['Compiled'] and not compilation['ProductionChanged']"
    assert code.count(anchor)==1
    code=code.replace(anchor,anchor+"\n    assert compilation['Modules']==['CCoupledDeformationContinuation']\n    compilation['Modules']=json.loads((root/'Preparation.json').read_bytes())['modules']",1)
    (root/'CompileClosure.json').write_bytes(closure)
    namespace={'__name__':'point_trace_source_composition_test','__file__':str(source_file)}
    exec(compile(code,str(source_file),'exec'),namespace)
    out=root/'PointTraceSourceComposition.json'
    namespace['TestAudit'](repo,root,out)
    audit=json.loads(out.read_bytes())
    audit['source_components_checked']=audit.pop('imported')
    audit['imported_this_pass']=['CCoupledDeformationContinuation']
    audit['other_actual_components_exact_parent_except_identifier_case']=unchanged
    audit['raw_source_case_changed_by_VBE']=different
    audit['raw_source_byte_identical_other_components']=len(unchanged)-len(different)
    audit['parent_compiled_project_sha256']=sha256(before).hexdigest().upper()
    out.write_bytes((json.dumps(audit,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode())
    (root/'ParentCompiledVBA.bin').write_bytes(before)
    prep=json.loads((root/'Preparation.json').read_bytes())
    prep['directed_scope'].update(single_test_module_compiled=False,
        actual_other_252_components_exact_except_identifier_case=True,
        actual_other_byte_identical_components=len(unchanged)-len(different),
        current_compiled_workbook_sha256=compilation['WorkbookSHA256'])
    for name in (Path(__file__).name,):
        (root/name).write_bytes(Path(__file__).read_bytes())
        prep['inputs'][name]=sha256((root/name).read_bytes()).hexdigest().upper()
    (root/'Preparation.json').write_bytes((json.dumps(prep,indent=2,sort_keys=True)+'\n').encode())
    print('STOCK YIELD CONTROLLER SOURCE VERIFIED; ONE CONTROLLER IMPORT, 252 PARENT COMPONENTS EXACT EXCEPT VBE IDENTIFIER CASE')


if __name__=='__main__':
    parser=ArgumentParser(description=__doc__);parser.add_argument('--repo',type=Path,default=Path.cwd())
    parser.add_argument('--root',type=Path,required=True);args=parser.parse_args()
    TestFinalize(args.repo.resolve(),args.root.resolve())
