# Переносит только тяжёлые данные текущего проекта на D с проверкой SHA-256 и сохранением логических путей.
# Фазы разделены: проверенная копия, переключение каталогов, удаление исключительно проверенных копий на C.
param([Parameter(Mandatory=$true)][ValidateSet('Prepare','Switch','Release')][string]$Phase)
$ErrorActionPreference='Stop'
$taskWorkspace=[IO.Path]::GetFullPath('C:\Users\avidzba\Desktop\программы\В_РАБОТЕ!\RC_Section_NDM_TQ')
$taskStorage=[IO.Path]::GetFullPath('D:\CodexScratch\RC_Section_NDM_TQ\WorkspaceData_20261127_CDrive')
$taskAudit=Join-Path $taskWorkspace 'tmp\disk-cleanup\20261127-CDrive'
$taskPlanPath=Join-Path $taskAudit 'Plan.json'
$taskBackupRoot=Join-Path $taskAudit 'VerifiedCBackups'

# Запрещает операции вне явно выбранных корней; проверяет абсолютный путь до перемещения или удаления.
function Assert-TaskPath([string]$Path,[string]$Root){
    $resolved=[IO.Path]::GetFullPath($Path)
    if(!$resolved.StartsWith($Root.TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw ('Путь вне разрешённого корня: '+$resolved)}
    return $resolved
}

# Считает контрольные суммы обычных файлов, не следует по вложенным junction и сохраняет их назначения отдельно.
function Get-TaskManifest([string]$Root){
    $files=[Collections.Generic.List[object]]::new();$links=[Collections.Generic.List[object]]::new()
    $stack=[Collections.Generic.Stack[IO.DirectoryInfo]]::new();$stack.Push([IO.DirectoryInfo]$Root);$bytes=0L
    while($stack.Count){
        $directory=$stack.Pop()
        foreach($entry in $directory.EnumerateFileSystemInfos()){
            $relative=[IO.Path]::GetRelativePath($Root,$entry.FullName)
            if(($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0){
                if(($entry.Attributes -band [IO.FileAttributes]::Directory) -eq 0){throw ('Неожиданная файловая ссылка: '+$entry.FullName)}
                $links.Add([ordered]@{Relative=$relative;Target=$entry.LinkTarget});continue
            }
            if(($entry.Attributes -band [IO.FileAttributes]::Directory) -ne 0){$stack.Push([IO.DirectoryInfo]$entry)}
            else{
                $size=([IO.FileInfo]$entry).Length;$hash=(Get-FileHash -LiteralPath $entry.FullName -Algorithm SHA256).Hash
                $files.Add([ordered]@{Relative=$relative;Bytes=$size;SHA256=$hash});$bytes+=$size
            }
        }
    }
    return [ordered]@{Files=$files.ToArray();Links=$links.ToArray();Bytes=$bytes}
}

# Проверяет полный набор файлов и ссылок, а не только размер каталога или отдельные образцы.
function Assert-TaskManifest($Expected,$Actual,[string]$Label){
    if($Expected.Files.Count -ne $Actual.Files.Count -or $Expected.Bytes -ne $Actual.Bytes -or $Expected.Links.Count -ne $Actual.Links.Count){throw ('Изменился состав '+$Label)}
    $map=@{};foreach($file in $Actual.Files){$map[$file.Relative]=$file}
    foreach($file in $Expected.Files){$other=$map[$file.Relative];if(!$other -or $other.Bytes -ne $file.Bytes -or $other.SHA256 -cne $file.SHA256){throw ('Не совпадает SHA '+$Label+'\'+$file.Relative)}}
    $map=@{};foreach($link in $Actual.Links){$map[$link.Relative]=$link.Target}
    foreach($link in $Expected.Links){if($map[$link.Relative] -cne $link.Target){throw ('Не совпадает ссылка '+$Label+'\'+$link.Relative)}}
}

# Сохраняет журнал с обеих сторон переноса; копия на D не зависит от будущего освобождения C.
function Save-TaskPlan($Plan){
    $raw=$Plan|ConvertTo-Json -Depth 12
    $raw|Set-Content -LiteralPath $taskPlanPath -Encoding utf8
    $raw|Set-Content -LiteralPath (Join-Path $taskStorage 'MoveManifest.json') -Encoding utf8
}

# Проверяет целостность Git по доступности объектов и неизменности HEAD и пользовательских изменений.
function Assert-TaskGit($Plan){
    $head=(& git -C $taskWorkspace rev-parse HEAD).Trim();if($LASTEXITCODE -ne 0 -or $head -cne $Plan.GitHEAD){throw 'Изменилась контрольная точка Git.'}
    $status=@(& git -C $taskWorkspace status --porcelain=v1 --untracked-files=normal)
    if($LASTEXITCODE -ne 0 -or (($status -join "`n") -cne ($Plan.GitStatus -join "`n"))){throw 'Изменилось состояние рабочей папки Git.'}
}

if($Phase -eq 'Prepare'){
    if((Test-Path -LiteralPath $taskPlanPath) -or (Test-Path -LiteralPath $taskStorage)){throw 'Каталог или план переноса уже существует.'}
    if(Get-Process -Name git -ErrorAction SilentlyContinue){throw 'Есть активный процесс Git; переключение данных запрещено.'}
    if((Get-PSDrive D).Free -lt 20GB){throw 'На D недостаточно запаса.'}
    $head=(& git -C $taskWorkspace rev-parse HEAD).Trim();if($LASTEXITCODE -ne 0){throw 'Не читается HEAD.'}
    $status=@(& git -C $taskWorkspace status --porcelain=v1 --untracked-files=normal)
    $plan=[ordered]@{Workspace=$taskWorkspace;Storage=$taskStorage;GitHEAD=$head;GitStatus=$status;Stage='Preparing';
        InitialCFreeBytes=(Get-PSDrive C).Free;InitialDFreeBytes=(Get-PSDrive D).Free;UTC=[DateTimeOffset]::UtcNow.ToString('o');Items=@()}
    New-Item -ItemType Directory -Path $taskStorage|Out-Null
    foreach($spec in @(@('GitObjects','.git\objects'),@('Evidence','docs\reports\evidence'),@('Literature','literature'))){
        $source=Assert-TaskPath (Join-Path $taskWorkspace $spec[1]) $taskWorkspace
        $destination=Assert-TaskPath (Join-Path $taskStorage $spec[0]) $taskStorage
        $backup=Assert-TaskPath (Join-Path $taskBackupRoot $spec[0]) $taskWorkspace
        if((Get-Item -LiteralPath $source -Force).Attributes -band [IO.FileAttributes]::ReparsePoint){throw ('Источник уже является ссылкой: '+$source)}
        $original=Get-TaskManifest $source
        Write-Output ('SOURCE_HASHED '+$spec[0]+' files='+$original.Files.Count+' bytes='+$original.Bytes)
        & robocopy.exe $source $destination /E /COPY:DAT /DCOPY:DAT /R:1 /W:1 /XJ /NFL /NDL /NP /NJH /NJS
        if($LASTEXITCODE -ge 8){throw ('Ошибка копирования '+$spec[0]+': '+$LASTEXITCODE)}
        foreach($link in $original.Links){
            $newLink=Assert-TaskPath (Join-Path $destination $link.Relative) $taskStorage
            New-Item -ItemType Directory -Path (Split-Path -Parent $newLink) -Force|Out-Null
            New-Item -ItemType Junction -Path $newLink -Target $link.Target|Out-Null
        }
        Assert-TaskManifest $original (Get-TaskManifest $destination) $spec[0]
        $plan.Items+=,[ordered]@{Name=$spec[0];Source=$source;Destination=$destination;Backup=$backup;Manifest=$original;State='CopiedAndHashed'}
        Save-TaskPlan $plan
        Write-Output ('COPY_SHA256_VERIFIED '+$spec[0])
    }
    Assert-TaskGit $plan;$plan.Stage='CopiesVerified';Save-TaskPlan $plan
    Write-Output ('PREPARE_COMPLETE bytes='+($plan.Items.Manifest.Bytes|Measure-Object -Sum).Sum)
    exit 0
}

if(!(Test-Path -LiteralPath $taskPlanPath)){throw 'Нет плана с проверенными копиями.'}
$plan=Get-Content -LiteralPath $taskPlanPath -Raw|ConvertFrom-Json -AsHashtable
if($plan.Workspace -cne $taskWorkspace -or $plan.Storage -cne $taskStorage){throw 'Изменились закреплённые пути.'}
if($Phase -eq 'Switch'){
    if($plan.Stage -cne 'CopiesVerified'){throw 'Копии ещё не проверены или переключение уже выполнялось.'}
    Assert-TaskGit $plan
    if(Get-Process -Name git -ErrorAction SilentlyContinue){throw 'Есть активный Git; переключение запрещено.'}
    foreach($item in $plan.Items){
        Assert-TaskManifest $item.Manifest (Get-TaskManifest $item.Source) ($item.Name+' source')
        Assert-TaskManifest $item.Manifest (Get-TaskManifest $item.Destination) ($item.Name+' destination')
    }
    New-Item -ItemType Directory -Path $taskBackupRoot|Out-Null
    foreach($item in $plan.Items){
        $null=Assert-TaskPath $item.Source $taskWorkspace;$null=Assert-TaskPath $item.Backup $taskWorkspace;$null=Assert-TaskPath $item.Destination $taskStorage
        if(Test-Path -LiteralPath $item.Backup){throw 'Резервный каталог уже существует.'}
        Move-Item -LiteralPath $item.Source -Destination $item.Backup
        try{New-Item -ItemType Junction -Path $item.Source -Target $item.Destination|Out-Null}
        catch{Move-Item -LiteralPath $item.Backup -Destination $item.Source;throw}
        $link=Get-Item -LiteralPath $item.Source -Force
        if(!($link.Attributes -band [IO.FileAttributes]::ReparsePoint) -or $link.LinkTarget -cne $item.Destination){throw ('Не подтверждён новый путь '+$item.Name)}
        $item.State='JunctionVerified_CBackupRetained';Save-TaskPlan $plan
        Write-Output ('JUNCTION_VERIFIED '+$item.Name)
    }
    Assert-TaskGit $plan
    $fsck=@(& git -C $taskWorkspace fsck --connectivity-only --no-dangling 2>&1);$fsckExit=$LASTEXITCODE
    $fsck|Set-Content -LiteralPath (Join-Path $taskStorage 'GitConnectivity.log') -Encoding utf8
    if($fsckExit -ne 0){throw ('Git fsck не прошёл: '+$fsckExit)}
    $plan.Stage='JunctionsAndGitVerified';$plan.GitConnectivityExitCode=$fsckExit;Save-TaskPlan $plan
    Write-Output 'SWITCH_COMPLETE; GIT CONNECTIVITY PASS; ALL C BACKUPS RETAINED'
    exit 0
}

if($plan.Stage -cne 'JunctionsAndGitVerified'){throw 'Удаление запрещено: нет успешной проверки ссылок и Git.'}
Assert-TaskGit $plan
foreach($item in $plan.Items){
    $backup=Assert-TaskPath $item.Backup $taskWorkspace;$null=Assert-TaskPath $item.Destination $taskStorage
    if((Get-Item -LiteralPath $backup -Force).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Резервный каталог не должен быть ссылкой.'}
    Assert-TaskManifest $item.Manifest (Get-TaskManifest $backup) ($item.Name+' C backup')
    Assert-TaskManifest $item.Manifest (Get-TaskManifest $item.Destination) ($item.Name+' D verified copy')
    foreach($link in $item.Manifest.Links){
        $linkPath=Assert-TaskPath (Join-Path $backup $link.Relative) $taskWorkspace
        $entry=Get-Item -LiteralPath $linkPath -Force
        if(!($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -or $entry.LinkTarget -cne $link.Target){throw 'Изменилась вложенная ссылка.'}
        [IO.Directory]::Delete($linkPath,$false) # Удаляется только запись junction внутри резервной копии; её данные на D остаются.
    }
    $remaining=Get-TaskManifest $backup;if($remaining.Links.Count){throw 'В резервной копии остались ссылки; рекурсивное удаление запрещено.'}
    Remove-Item -LiteralPath $backup -Recurse -Force
    $item.State='StoredOnD_CCopyRemoved';Save-TaskPlan $plan
    Write-Output ('C_COPY_REMOVED_AFTER_FULL_SHA256_CHECK '+$item.Name)
}
Assert-TaskGit $plan
$plan.Stage='Complete';$plan.FinalCFreeBytes=(Get-PSDrive C).Free;$plan.FinalDFreeBytes=(Get-PSDrive D).Free
$plan.FreedCBytes=$plan.FinalCFreeBytes-$plan.InitialCFreeBytes;$plan.CompletedUTC=[DateTimeOffset]::UtcNow.ToString('o')
Save-TaskPlan $plan
Write-Output ('COMPLETE C_FREED_GIB='+[math]::Round($plan.FreedCBytes/1GB,3)+' C_FREE_GIB='+[math]::Round($plan.FinalCFreeBytes/1GB,3))
