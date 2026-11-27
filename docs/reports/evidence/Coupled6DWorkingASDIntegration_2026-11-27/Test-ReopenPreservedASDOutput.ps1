# ДЛЯ ТЕСТОВ. Компилирует повторно открытую сохранную книгу только для чтения и экспортирует фактический VBA без Save.
param([Parameter(Mandatory=$true)][string]$Root)
$ErrorActionPreference='Stop'
Push-Location -LiteralPath $Root
try {
    . './Guard-IsolatedOffice.ps1'
    $proof=Get-Content -LiteralPath 'PreservedPackageProof.json' -Raw | ConvertFrom-Json
    $path=(Resolve-Path -LiteralPath 'docs/regression/Diagnostic.xlsm').Path
    if((Get-FileHash -LiteralPath $path).Hash -cne $proof.saved_candidate_sha256){throw 'Preserved candidate changed.'}
    Assert-IsolatedOfficeReady 'PreservedOutputReadOnlyCompileAndDump'
    Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class PreservedASDWindowTest { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId); }'
    $excel=$null;$book=$null;$taskCloseAnchor=$null;$ownedPid=0;$ownedTicks=0
    $preexisting=@(Get-Process Excel -ErrorAction SilentlyContinue)
    try {
        $excel=New-Object -ComObject Excel.Application
        [uint32]$taskOfficeId=0;[void][PreservedASDWindowTest]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$taskOfficeId)
        if($taskOfficeId -le 0 -or $taskOfficeId -in $preexisting.Id){throw 'Excel is not a newly created owned process.'}
        $worker=Get-Process -Id $taskOfficeId -ErrorAction Stop
        if($worker.ProcessName -ine 'EXCEL'){throw 'Invalid Excel owner.'}
        $ownedPid=[int]$taskOfficeId;$ownedTicks=$worker.StartTime.ToUniversalTime().Ticks
        Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
        $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
        $book=$excel.Workbooks.Open($path,0,$true)
        if(!$book.ReadOnly){throw 'Candidate must be read-only.'}
        $project=$book.VBProject;$components=$project.VBComponents
        if($components.Count -ne 255){throw 'Unexpected component count.'}
        $excel.VBE.ActiveVBProject=$project
        $compile=$excel.VBE.CommandBars.FindControl(1,578)
        if(!$compile){throw 'Compile command unavailable.'}
        $initiallyEnabled=[bool]$compile.Enabled
        if($compile.Enabled){$compile.Execute()}
        for($probe=0;$probe -lt 40 -and $compile.Enabled;$probe++){Start-Sleep -Milliseconds 250}
        if($compile.Enabled){throw 'Preserved candidate failed native compilation.'}
        $writer=[IO.StreamWriter]::new((Join-Path (Split-Path -Parent $path) 'VBA_All_Code.txt'),$false,[Text.UTF8Encoding]::new($true))
        try {
            $writer.WriteLine('VBA PROJECT EXPORT');$writer.WriteLine('Workbook: RC_Section_NDM.xlsm');$writer.WriteLine('Date: '+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
            foreach($component in $components){
                $module=$component.CodeModule
                $writer.WriteLine('='*100);$writer.WriteLine('COMPONENT: '+$component.Name);$writer.WriteLine('TYPE: '+$component.Type);$writer.WriteLine('LINES: '+$module.CountOfLines);$writer.WriteLine('='*100)
                if($module.CountOfLines){$writer.WriteLine($module.Lines(1,$module.CountOfLines))}
                $writer.WriteLine()
            }
        } finally {$writer.Dispose()}
        [ordered]@{SavedReopened=$true;ReadOnly=$book.ReadOnly;Components=$components.Count;Compiled=$true;CompileCommandInitiallyEnabled=$initiallyEnabled;VbaDumpFromActualReopenedBook=$true;NoSave=$true} |
            ConvertTo-Json | Set-Content -LiteralPath 'Native/PreservedReopenProof.json' -Encoding UTF8
        $taskCloseAnchor=$excel.Workbooks.Add();$taskCloseAnchor.Saved=$true
        foreach($taskVbeObject in @($compile,$module,$component,$components,$project)){
            if($null -ne $taskVbeObject -and [Runtime.InteropServices.Marshal]::IsComObject($taskVbeObject)){
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskVbeObject)
            }
        }
        $taskVbeObject=$null;$compile=$null;$module=$null;$component=$null;$components=$null;$project=$null
        $book.Close($false);$book=$null
        if((Get-FileHash -LiteralPath $path).Hash -cne $proof.saved_candidate_sha256){throw 'Read-only verification changed candidate.'}
        Write-Output 'PRESERVED_OUTPUT_REOPEN_COMPILE_DUMP_PASS;SOURCE_UNCHANGED=True'
    } finally {
        if($book){try{$book.Close($false)}catch{Write-Output ('OWN_CLOSE_WARNING: '+$_.Exception.Message)}}
        if($excel -and $ownedPid){try{$excel.Quit()}catch{Write-Output ('OWN_QUIT_WARNING: '+$_.Exception.Message)}}
        if($ownedPid){
            $worker=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
            if($worker -and $worker.ProcessName -eq 'EXCEL' -and $worker.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){Stop-Process -Id $ownedPid -Force}
        }
        if($excel){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}catch{}}
        $excel=$null;$book=$null;$taskCloseAnchor=$null
        [GC]::Collect()
    }
} finally {Pop-Location}
