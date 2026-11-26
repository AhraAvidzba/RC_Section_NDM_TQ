# ДЛЯ ТЕСТОВ. Держит только собственную видимую копию книги для нативной проверки одного списка CDP/ASD.
param([Parameter(Mandatory=$true)][string]$Root)
$ErrorActionPreference='Stop'
Push-Location -LiteralPath $Root
$application=$null; $book=$null; $ownedId=0; $ownedTicks=0
try {
    . './Guard-IsolatedOffice.ps1'
    Assert-IsolatedOfficeReady 'NativeSelectableMaterialDropdown'
    $compiled=Get-Content -LiteralPath 'docs/regression/CodeUpdate.json' -Raw | ConvertFrom-Json
    $source='docs/regression/Diagnostic.xlsm'
    if(!$compiled.Compiled -or (Get-FileHash -LiteralPath $source).Hash -cne $compiled.WorkbookSHA256){throw 'Compiled source book mismatch'}
    $uiRoot=Join-Path $Root 'Native/UI'; if(Test-Path -LiteralPath $uiRoot){throw 'UI session already exists'}
    New-Item -ItemType Directory -Path $uiRoot | Out-Null
    $copy=Join-Path $uiRoot 'SelectableLaws-UI.xlsm';Copy-Item -LiteralPath $source -Destination $copy
    if(-not ('SelectableLawWindowTest' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class SelectableLawWindowTest {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);
}
'@
    }
    $existing=@(Get-Process Excel -ErrorAction SilentlyContinue | ForEach-Object {$_.Id})
    $application=New-Object -ComObject Excel.Application
    [uint32]$identifier=0;[void][SelectableLawWindowTest]::GetWindowThreadProcessId([IntPtr]$application.Hwnd,[ref]$identifier)
    if($identifier -le 0 -or $identifier -in $existing){throw 'Excel owner not new'}
    $ownedId=[int]$identifier;$ownedTicks=(Get-Process -Id $ownedId).StartTime.ToUniversalTime().Ticks
    [ordered]@{PID=$ownedId;StartTicksUTC=[string]$ownedTicks;SourceSHA256=$compiled.WorkbookSHA256;Copy=$copy} |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $uiRoot 'Ownership.json') -Encoding UTF8
    $application.EnableEvents=$false
    $book=$application.Workbooks.Open($copy,0,$false)
    $table=$book.Names.Item('rngCoupledMaterialSettings').RefersToRange
    $cell=$null
    for($r=2;$r -le $table.Rows.Count;$r++) {
        if([string]$table.Cells.Item($r,1).Value2 -ceq 'Coupled.ConcreteLaw'){$cell=$table.Cells.Item($r,2);break}
    }
    if($null -eq $cell){throw 'Selector key not found'}
    $application.Visible=$true;$cell.Worksheet.Activate();$application.Goto($cell,$true)
    $deadline=[DateTime]::UtcNow.AddMinutes(20);$sequence=0;$stop=$false
    do {
        $commandPath=Join-Path $uiRoot 'Command.json'
        if(Test-Path -LiteralPath $commandPath) {
            $command=Get-Content -LiteralPath $commandPath -Raw | ConvertFrom-Json
            if([int]$command.Sequence -gt $sequence) {
                $sequence=[int]$command.Sequence
                switch([string]$command.Action) {
                    'Status' { }
                    'SaveReopen' {
                        $book.Save();$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
                        $book=$application.Workbooks.Open($copy,0,$false)
                        $table=$book.Names.Item('rngCoupledMaterialSettings').RefersToRange
                        $cell=$table.Cells.Item($r,2);$cell.Worksheet.Activate();$application.Goto($cell,$true)
                    }
                    'Stop' {$stop=$true}
                    default {throw 'Unsupported UI test command'}
                }
                $law=$application.Run("'"+$book.Name+"'!modTestSelectableMaterialCore.TestReadSelectableConcreteLaw")
                [ordered]@{Sequence=$sequence;Action=$command.Action;Value=[string]$cell.Value2;ValidationAcceptsCurrent=[bool]$cell.Validation.Value;
                    Formula1=[string]$cell.Validation.Formula1;Address=[string]$cell.Address();ReaderLaw=[string]$law;
                    Saved=[bool]$book.Saved;PID=$ownedId;StartTicksUTC=[string]$ownedTicks} |
                    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $uiRoot ('Status'+$sequence+'.json')) -Encoding UTF8
            }
        }
        Start-Sleep -Milliseconds 200
    } while(!$stop -and [DateTime]::UtcNow -lt $deadline)
    if(!$stop){throw 'UI test session time limit exceeded'}
} finally {
    if($book){$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
    if($application -and $ownedId -gt 0) {
        $worker=Get-Process -Id $ownedId -ErrorAction SilentlyContinue
        if($worker -and $worker.ProcessName -ieq 'EXCEL' -and $worker.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){$application.Quit()}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
    }
    Pop-Location
}
