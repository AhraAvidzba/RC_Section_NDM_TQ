"""ДЛЯ ТЕСТОВ. Готовит полную отдельную сборку из рабочих исходников с переключателем CDP/ASD, не заменяя релиз."""
from pathlib import Path
from hashlib import sha256
import shutil,json

repo=Path.cwd();base=Path('D:/CodexScratch/RC_Section_NDM_TQ');root=base/'Cold1017SelectableMaterialBook_20261126';assert not root.exists()
root.mkdir()
for name in ('src','tests'):shutil.copytree(repo/name,root/name)
shutil.copytree(repo/'tools/build_workbook',root/'tools/build_workbook')
for name in ('Temp','Native','docs/geometry_refactor','workbook/output'):(root/name).mkdir(parents=True,exist_ok=True)
layout='docs/geometry_refactor/SECTION_RECIPE_GUIDE_LAYOUT.json';shutil.copyfile(repo/layout,root/layout)
guide=json.loads((repo/layout).read_bytes())
for name,digest in guide['sourceHashes'].items():
    assert sha256((repo/name).read_bytes()).hexdigest().lower()==digest.lower()
    target=root/name;target.parent.mkdir(parents=True,exist_ok=True)
    if not target.exists():shutil.copyfile(repo/name,target)
for name in ('docs/NDM_User_Guide.md','docs/Coupled6DMaterialContract.md'):
    shutil.copyfile(repo/name,root/name)
images=repo/'workbook/output/formula_images'
if images.exists():shutil.copytree(images,root/'workbook/output/formula_images')
parent=base/'Cold1016PureBendingSteelLimitFixed_20261126'
shutil.copyfile(parent/'Guard-IsolatedOffice.ps1',root/'Guard-IsolatedOffice.ps1')
shutil.copyfile(Path(__file__),root/Path(__file__).name)
runner='''# ДЛЯ ТЕСТОВ. Собирает отдельную книгу из закреплённых рабочих исходников; релиз не заменяет.
param([Parameter(Mandatory=$true)][string]$Root)
$ErrorActionPreference='Stop'
Push-Location -LiteralPath $Root
$previousTemp=$env:TEMP; $previousTmp=$env:TMP
try {
    . './Guard-IsolatedOffice.ps1'
    $prep=Get-Content -LiteralPath 'Preparation.json' -Raw | ConvertFrom-Json
    foreach($pin in $prep.inputs.psobject.Properties) {
        if((Get-FileHash -LiteralPath $pin.Name).Hash -cne $pin.Value){throw ('Changed build input '+$pin.Name)}
    }
    Assert-IsolatedOfficeReady 'BuildSelectableMaterialBook'
    $env:TEMP=(Resolve-Path -LiteralPath 'Temp').Path; $env:TMP=$env:TEMP
    & './tools/build_workbook/Build-Workbook.ps1' -OutputPath 'workbook/output/SelectableLaws.xlsm' -UseFormattedLayout |
        Tee-Object -FilePath 'Build.log'
    $owner=Get-Content -LiteralPath 'workbook/output/BuildOwnership.json' -Raw | ConvertFrom-Json
    $deadline=[DateTime]::UtcNow.AddSeconds(20)
    do {
        $worker=Get-Process -Id $owner.PID -ErrorAction SilentlyContinue
        if(!$worker){break}
        if($worker.ProcessName -ine 'EXCEL' -or $worker.StartTime.ToUniversalTime().Ticks -ne [long]$owner.StartTicksUTC){throw 'Build process identity changed'}
        Start-Sleep -Milliseconds 200
    } while([DateTime]::UtcNow -lt $deadline)
    if($worker){throw 'Own build process has not closed'}
    [ordered]@{State='Absent';PID=$owner.PID;StartTicksUTC=$owner.StartTicksUTC;SavedWorkbookSHA256=(Get-FileHash -LiteralPath 'workbook/output/SelectableLaws.xlsm').Hash} |
        ConvertTo-Json | Set-Content -LiteralPath 'BuildClosure.json' -Encoding UTF8
} finally {
    $env:TEMP=$previousTemp; $env:TMP=$previousTmp
    Pop-Location
}
'''
(root/'Run-TestSelectableMaterialBook.ps1').write_bytes(runner.replace('\n','\r\n').encode('utf-8-sig'))
prep=dict(case='FreshSelectableMaterialBook',full_acceptance_not_run=True,source_edits_not_yet_validated=True,
    inputs={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*') if p.is_file()},
    release_hashes={name:sha256((repo/name).read_bytes()).hexdigest().upper() for name in ('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt')})
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREPARED1017; COMPLETE WORKING SOURCE BUILD; CDP/ASD SELECTOR; RELEASE ISOLATED; NO ACCEPTANCE CLAIM')
