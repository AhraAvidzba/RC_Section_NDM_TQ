"""ДЛЯ ТЕСТОВ. Расширяет существующий линейно отображённый резерв ASD на смешанную меру в отдельной книге."""
from pathlib import Path
from hashlib import sha256
import json
import shutil

repo=Path.cwd()
base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold1153ASDPhysicalDirections_20261127'
root=base/'Cold1154ASDLinearTrustMy_20261127'
assert not root.exists()
for folder in ['src/Solver','tests','Input','Native','Temp','Before','docs/regression','workbook/output']:
    (root/folder).mkdir(parents=True,exist_ok=True)
for name in ['tools','tests']:
    shutil.copytree(parent/name,root/name,dirs_exist_ok=True)
for name in ['Guard-IsolatedOffice.ps1','Run-TestASDPhysicalDirections.ps1']:
    shutil.copyfile(parent/name,root/name)
book=(repo/'workbook/output/RC_Section_NDM.xlsm').read_bytes()
assert sha256(book).hexdigest().upper()=='1E6443C2CFDD0058FE5EA559D1D45A8E94CC5199A3ADE69149A06CD4265D8559'
for name in ['Before/RC_Section_NDM.xlsm','docs/regression/Diagnostic.xlsm','workbook/output/RC_Section_NDM.xlsm']:
    (root/name).write_bytes(book)
relative='src/Solver/CCoupledNewtonSolver.cls'
source=(repo/relative).read_text('utf-8-sig')
start=source.index("' Выбирает точное линейное отображение только для одной канонической кривизны или закручивания ASD.")
end=source.index('\nEnd Function',start)+len('\nEnd Function')
old=source[start:end]
new='''\' Выбирает точное линейное отображение заданной меры ASD, включая осевую и смешанную пространственную проекции.
\' Нелинейная материальная координата и псевдодуга сохраняют собственные уравнения и правила восстановления.
Private Function HasASDLinearControl() As Boolean
    If Not HasPrincipalBoundaries Or Not mControlActive Or mEnvelopeControl Or mArcLoadProjection <> 0# Then Exit Function
    Dim degree As Long
    For degree = 1 To mGaugeStart - 1
        If mControlProjection(degree) <> 0# Then HasASDLinearControl = True: Exit Function
    Next degree
End Function'''
assert source.count('HasASDGradientControl()')==3
source=source[:start]+new+source[end:]
source=source.replace('HasASDGradientControl()','HasASDLinearControl()')
(root/relative).write_text(source,encoding='utf-8-sig',newline='\r\n')
(root/'Before/CCoupledNewtonSolver.cls').write_bytes((repo/relative).read_bytes())
compiler=(parent/'Compile-TestASDPhysicalDirections.ps1').read_text('utf-8-sig')
compiler=compiler.replace("@('modTestSelectableMaterialCore','CCoupledMaterialProgressTest')",
                          "@('CCoupledNewtonSolver','modTestSelectableMaterialCore','CCoupledMaterialProgressTest')",1)
(root/'Compile-TestASDLinearTrustMy.ps1').write_text(compiler,encoding='utf-8-sig',newline='\r\n')
for name in ['Case01','Case02','Case03','Case04']:
    case=root/name
    for folder in ['Input','Native','docs/regression']:
        (case/folder).mkdir(parents=True,exist_ok=True)
    shutil.copytree(parent/name/'Input',case/'Input',dirs_exist_ok=True)
    for relative in ['Case.json','Native/MainCapacity.Operators.csv']:
        shutil.copyfile(parent/name/relative,case/relative)
(root/'ParentCodeUpdate.json').write_text(json.dumps(dict(WorkbookSHA256=sha256(book).hexdigest().upper()))+'\n',encoding='utf-8')
pins={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper()
      for p in root.rglob('*') if p.is_file() and p.relative_to(root).as_posix()!='docs/regression/Diagnostic.xlsm'}
(root/'Preparation.json').write_text(json.dumps(dict(inputs=pins,base_git_commit='c816ed3',
       public_book_sha256=sha256(book).hexdigest().upper(),changed_kernel_components=['CCoupledNewtonSolver'],
       physical_model_and_tolerances_unchanged=True,expected_components=262,
       only_change='Existing QR trust-region least-squares and exact linear mapping also used for ASD axial and mixed field controls'),
       sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('ISOLATED ASD LINEAR TRUST CANDIDATE STAGED; ONE KERNEL CLASS; NO PHYSICAL MODEL OR TOLERANCE CHANGE')
