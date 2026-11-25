"""ДЛЯ ТЕСТОВ. Готовит строго ограниченное восстановление паспорта после RPC-ошибки Close уже сохранённой частной книги."""
from pathlib import Path

here=Path(__file__).resolve().parent
root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold837ElasticPrincipalEvent_20261125')
code=(here/'Recover-TestSavedCompilation.ps1').read_text('utf-8-sig')
old="if($failure.Message -ne 'You cannot call a method on a null-valued expression.' -or $failure.ScriptStackTrace -notmatch 'line 201'){throw 'Другая ошибка компиляции.'}"
assert code.count(old)==1
code=code.replace(old,"if($failure.Message -notmatch '0x80010100' -or $failure.ScriptStackTrace -notmatch 'line 212'){throw 'Другая ошибка компиляции.'}")
old="if($source[200].Trim() -cne '$taskCloseAnchor=$excel.Workbooks.Add()'){throw 'Ошибка не относится к созданию служебной книги после Save.'}"
assert code.count(old)==1
code=code.replace(old,"if($source[211].Trim() -cne '$book.Close($false);$book=$null'){throw 'Ошибка не относится к Close после Save.'}")
code=code.replace("    $original=Get-Content -LiteralPath 'ParentCodeUpdate.json'", "    $prep=Get-Content -LiteralPath 'Preparation.json' -Raw|ConvertFrom-Json\n    $original=Get-Content -LiteralPath 'ParentCodeUpdate.json'",1)
code=code.replace("Modules=@('modTestCoupledElasticState');",'Modules=@($prep.directed_scope.actual_import_this_pass);',1)
code=code.replace('MandatoryActual255SourceAuditStillRequired','MandatoryActual256SourceAuditStillRequired')
(root/'Recover-TestSavedCloseCompilation.ps1').write_bytes(code.replace('\n','\r\n').encode('utf-8-sig'))
print(root/'Recover-TestSavedCloseCompilation.ps1')
