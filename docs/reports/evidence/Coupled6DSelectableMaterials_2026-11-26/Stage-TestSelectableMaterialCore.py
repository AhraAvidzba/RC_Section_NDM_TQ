"""ДЛЯ ТЕСТОВ. Добавляет небольшой направленный тест в проверенную свежую книгу с рабочими классами CDP/ASD."""
from pathlib import Path
from hashlib import sha256
import json,shutil

here=Path(__file__).resolve().parent;repo=Path.cwd();base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold1017SelectableMaterialBook_20261126';root=base/'Cold1018SelectableMaterialCore_20261126';assert not root.exists()
proof=json.loads((parent/'FreshSourceProof.json').read_bytes());assert proof['changed_core_modules_strictly_match_pinned_sources']
book=(parent/'workbook/output/SelectableLaws.xlsm').read_bytes();assert sha256(book).hexdigest().upper()==proof['saved_book_sha256']
root.mkdir()
for directory in ('src','tests','tools'):shutil.copytree(parent/directory,root/directory)
for directory in ('Native','Temp','docs/regression','workbook/output'):(root/directory).mkdir(parents=True,exist_ok=True)
(root/'docs/regression/Diagnostic.xlsm').write_bytes(book)
shutil.copyfile(repo/'workbook/output/RC_Section_NDM.xlsm',root/'workbook/output/RC_Section_NDM.xlsm')
shutil.copyfile(parent/'Guard-IsolatedOffice.ps1',root/'Guard-IsolatedOffice.ps1')
legacy=base/'Cold1016PureBendingSteelLimitFixed_20261126'
shutil.copyfile(legacy/'tools/build_workbook/Update-IsolatedOptimizerCode.ps1',root/'tools/build_workbook/Update-IsolatedOptimizerCode.ps1')
module=(here/'TestSelectableMaterialCore.bas.txt').read_text('utf-8-sig')
(root/'tests/modTestSelectableMaterialCore.bas').write_bytes(module.replace('\n','\r\n').encode('utf-8-sig'))
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig');assert compiler.count("-ModuleNames 'modTestCoupledElasticState'")==1
(root/'Compile-TestSelectableMaterialCore.ps1').write_bytes(compiler.replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'modTestSelectableMaterialCore'").replace('\n','\r\n').encode('utf-8-sig'))
runner=(legacy/'Run-TestPureBendingCurvature.ps1').read_text('utf-8-sig').replace('modTestCoupledElasticState.TestRunPureBendingCurvature','modTestSelectableMaterialCore.TestRunSelectableMaterialCore')
(root/'Run-TestSelectableMaterialCore.ps1').write_bytes(runner.replace('\n','\r\n').encode('utf-8-sig'))
(root/'ParentCodeUpdate.json').write_text(json.dumps(dict(Compiled=True,ProductionChanged=False,WorkbookSHA256=proof['saved_book_sha256']),indent=2)+'\n',encoding='utf-8')
(root/'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent),fresh_source_proof=proof),indent=2)+'\n',encoding='utf-8')
shutil.copyfile(Path(__file__),root/Path(__file__).name)
prep=dict(directed_scope=dict(case='SelectableMaterialCoreSmoke',actual_import_this_pass=['modTestSelectableMaterialCore'],
    actual_component_count=256,all255_parent_components_retained=True,physical_capacity_not_tested=True),
    inputs={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
            if p.is_file() and p.relative_to(root).as_posix()!='docs/regression/Diagnostic.xlsm'})
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREPARED1018; EIGHT SMALL WORKING-CORE CASES; ACTUAL SELECTOR READER; EXTERNAL ASD MATERIAL REPLAY INPUT')
