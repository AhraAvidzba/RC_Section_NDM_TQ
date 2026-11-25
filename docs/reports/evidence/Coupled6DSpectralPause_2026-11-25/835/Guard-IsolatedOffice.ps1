$protected=@{16688=[long]639305194106046481;25900=[long]639301894503450644;32172=[long]639304180394977771;36308=[long]639304159664638157;37736=[long]639304161194652457}
$previousTemp=$env:TEMP;$previousTmp=$env:TMP
$previousExportRoot=$env:NDM_SECANT_EXPORT_ROOT;$previousSourceRoot=$env:NDM_LOAD_PATH_VERIFIER_ROOT

# Проверяет свободные ресурсы и защищает все подтверждённые пользовательские экземпляры.
function Assert-IsolatedOfficeReady([string]$Phase) {
    $others=@(Get-Process Excel -ErrorAction SilentlyContinue)
    # ДЛЯ ТЕСТОВ: все существующие экземпляры внешние для нового запуска.
    # Живой собственный компилятор исключается по точным PID и времени.
    $ownCompiler=@()
    if (Test-Path -LiteralPath 'Compile.log') {
        foreach($line in @(Get-Content -LiteralPath 'Compile.log' -Encoding utf8)) {
            if ($line -match '^OWNED_EXCEL: id=(\d+); utcStartTicks=(\d+)$') {
                $ownCompiler+=@{PID=[int]$Matches[1];Ticks=[long]$Matches[2]}
            }
        }
    }
    # ДЛЯ ТЕСТОВ: все существующие экземпляры внешние для нового запуска.
    # Живой собственный компилятор исключается по точным PID и времени.
    $ownCompiler=@()
    if (Test-Path -LiteralPath 'Compile.log') {
        foreach($line in @(Get-Content -LiteralPath 'Compile.log' -Encoding utf8)) {
            if ($line -match '^OWNED_EXCEL: id=(\d+); utcStartTicks=(\d+)$') {
                $ownCompiler+=@{PID=[int]$Matches[1];Ticks=[long]$Matches[2]}
            }
        }
    }
    foreach($item in $others) {
        $ticks=$item.StartTime.ToUniversalTime().Ticks
        foreach($own in $ownCompiler) {
            if ($item.Id -eq $own.PID -and $ticks -eq $own.Ticks) {throw 'Собственный компилятор ещё работает.'}
        }
        $protected[$item.Id]=$ticks
    }
    $unknown=@($others | Where-Object {!$protected.ContainsKey($_.Id) -or $protected[$_.Id] -ne $_.StartTime.ToUniversalTime().Ticks})
    $memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
    $state=[ordered]@{phase=$Phase;utc=[DateTimeOffset]::UtcNow.ToString('o');ramMiB=$memory.AvailableMBytes;
        freeCommitGiB=($memory.CommitLimit-$memory.CommittedBytes)/1GB;freeC_MiB=(Get-PSDrive C).Free/1MB;freeD_GiB=(Get-PSDrive D).Free/1GB;
        excel=@($others|Select-Object Id,StartTime,@{Name='ProtectedExisting';Expression={$protected[$_.Id] -eq $_.StartTime.ToUniversalTime().Ticks}})}
    $state|ConvertTo-Json -Depth 5 -Compress|Add-Content -LiteralPath "Run.ndjson" -Encoding utf8
    if ($unknown.Count -or $state.ramMiB -lt 6500 -or $state.freeCommitGiB -lt 6 -or $state.freeC_MiB -lt 80 -or $state.freeD_GiB -lt 4) {
        throw "Нет свободного окна для этапа $Phase. Подробности в Run.ndjson."
    }
}

# Подтверждает закрытие по зафиксированной личности VBA COM-процесса.
function Save-OfficeClosure([string]$Log,[string]$Output) {
    $pattern='^OWNED_EXCEL: id=(\d+); utcStartTicks=(\d+)$'
    $owner=@(Get-Content -LiteralPath $Log -Encoding utf8 | Where-Object {$_ -match $pattern})
    if ($owner.Count -ne 1 -or $owner[0] -notmatch $pattern) {throw "Личность Excel отсутствует в $Log."}
    $workerId=[int]$Matches[1];$workerTicks=[long]$Matches[2]
    if ($workerId -le 0 -or $workerTicks -le 0 -or ($protected.ContainsKey($workerId) -and $protected[$workerId] -eq $workerTicks)) {throw 'Недопустимая личность Excel.'}
    $deadline=[DateTime]::UtcNow.AddSeconds(20)
    do {
        $worker=Get-Process -Id $workerId -ErrorAction SilentlyContinue
        if (!$worker) {break}
        if ($worker.ProcessName -ine 'excel' -or $worker.StartTime.ToUniversalTime().Ticks -ne $workerTicks) {throw 'Личность дочернего Excel изменилась.'}
        Start-Sleep -Milliseconds 200
    } while ([DateTime]::UtcNow -lt $deadline)
    if ($worker) {throw 'Дочерний Excel ещё работает.'}
    [ordered]@{State='Absent';PID=$workerId;StartTicksUTC=[string]$workerTicks} |
        ConvertTo-Json | Set-Content -LiteralPath $Output -Encoding utf8
}

