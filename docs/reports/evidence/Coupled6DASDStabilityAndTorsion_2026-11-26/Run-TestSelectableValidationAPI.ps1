# ДЛЯ ТЕСТОВ. Проверяет принятие значений и сохранение списка через нативный Excel API; визуальную приёмку не заменяет.
param([Parameter(Mandatory=$true)][string]$Root)
$ErrorActionPreference='Stop'
Push-Location -LiteralPath $Root
$application=$null;$book=$null;$ownedId=0;$ownedTicks=0
function Get-TestLawCell([object]$Workbook) {
    $range=$Workbook.Names.Item('rngCoupledMaterialSettings').RefersToRange
    for($index=2;$index -le $range.Rows.Count;$index++) {
        if([string]$range.Cells.Item($index,1).Value2 -ceq 'Coupled.ConcreteLaw'){return $range.Cells.Item($index,2)}
    }
    throw 'Selector key missing'
}
try {
    . './Guard-IsolatedOffice.ps1';Assert-IsolatedOfficeReady 'SelectableNativeValidationAPI'
    $area=Join-Path $Root 'Native/ValidationAPI';if(Test-Path -LiteralPath $area){throw 'Validation API output already exists'}
    New-Item -ItemType Directory -Path $area | Out-Null
    $compiled=Get-Content -LiteralPath 'docs/regression/CodeUpdate.json' -Raw | ConvertFrom-Json
    if((Get-FileHash -LiteralPath 'docs/regression/Diagnostic.xlsm').Hash -cne $compiled.WorkbookSHA256){throw 'Compiled book mismatch'}
    $copy=Join-Path $area 'ValidationAPI.xlsm';Copy-Item -LiteralPath 'docs/regression/Diagnostic.xlsm' -Destination $copy
    if(-not ('SelectableNativeAPIWindowTest' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class SelectableNativeAPIWindowTest {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);
}
'@
    }
    $existing=@(Get-Process Excel -ErrorAction SilentlyContinue | ForEach-Object {$_.Id})
    $application=New-Object -ComObject Excel.Application
    [uint32]$identifier=0;[void][SelectableNativeAPIWindowTest]::GetWindowThreadProcessId([IntPtr]$application.Hwnd,[ref]$identifier)
    if($identifier -le 0 -or $identifier -in $existing){throw 'Own Excel identity not new'}
    $ownedId=[int]$identifier;$ownedTicks=(Get-Process -Id $ownedId).StartTime.ToUniversalTime().Ticks
    [ordered]@{PID=$ownedId;StartTicksUTC=[string]$ownedTicks} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $area 'Ownership.json') -Encoding UTF8
    $application.Visible=$false;$application.EnableEvents=$false;$application.DisplayAlerts=$false
    $book=$application.Workbooks.Open($copy,0,$false);$cell=Get-TestLawCell $book
    $original=[string]$cell.Value2;$valid=@();$invalid=@()
    if($cell.Validation.Type -ne 3 -or $cell.Validation.AlertStyle -ne 1 -or !$cell.Validation.ShowError){throw 'Validation is not an active Stop list'}
    $formula=[string]$cell.Validation.Formula1
    $format=[ordered]@{Color=[int]$cell.Interior.Color;HorizontalAlignment=[int]$cell.HorizontalAlignment;WrapText=[bool]$cell.WrapText}
    foreach($choice in @('CDP','ASD')) {
        $cell.Value2=$choice;if(!$cell.Validation.Value){throw ('Native validation rejects '+$choice)}
        $book.Save();$cell=$null;$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
        $book=$application.Workbooks.Open($copy,0,$false);$cell=Get-TestLawCell $book
        $law=[string]$application.Run("'"+$book.Name+"'!modTestSelectableMaterialCore.TestReadSelectableConcreteLaw")
        if([string]$cell.Value2 -cne $choice -or !$cell.Validation.Value -or $law -cne $choice){throw 'Saved selector/reader mismatch'}
        $valid+=,[ordered]@{Value=$choice;AcceptedAfterReopen=$true;ReaderLaw=$law;Formula1=[string]$cell.Validation.Formula1}
    }
    foreach($choice in @('CDP,ASD','CDP;ASD','UnknownLaw')) {
        $cell.Value2=$choice;$accepted=[bool]$cell.Validation.Value
        $invalid+=,[ordered]@{Value=$choice;NativeValidationAccepts=$accepted}
        if($accepted){throw ('Native validation accepts invalid value '+$choice)}
    }
    $cell.Value2=$original;$book.Save()
    [ordered]@{NativeAPIValuesPassed=$true;NativeDropdownUIPassed=$false;Formula1=$formula;Valid=$valid;Invalid=$invalid;InputFormat=$format;
        OriginalSelectorRestored=$true;OriginalValue=$original;SourceBookUnchanged=((Get-FileHash -LiteralPath 'docs/regression/Diagnostic.xlsm').Hash -ceq $compiled.WorkbookSHA256)} |
        ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $area 'ValidationAPIProof.json') -Encoding UTF8
} catch {
    if(Test-Path -LiteralPath $area) {
        [ordered]@{Message=$_.Exception.Message;Position=$_.InvocationInfo.PositionMessage;NativeDropdownUIPassed=$false} |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $area 'ValidationAPIFailure.json') -Encoding UTF8
    }
    throw
} finally {
    if($book){$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
    if($application -and $ownedId -gt 0){
        $worker=Get-Process -Id $ownedId -ErrorAction SilentlyContinue
        if($worker -and $worker.ProcessName -ieq 'EXCEL' -and $worker.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){$application.Quit()}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
    }
    Pop-Location
}
