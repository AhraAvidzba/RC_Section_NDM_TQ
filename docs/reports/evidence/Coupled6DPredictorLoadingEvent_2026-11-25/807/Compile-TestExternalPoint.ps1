# ДЛЯ ТЕСТОВ: импортирует один читающий тестовый модуль в проверенную частную книгу.
$ErrorActionPreference='Stop'
Push-Location -LiteralPath $PSScriptRoot
try {
    . './Guard-IsolatedOffice.ps1'
    $prep=Get-Content -LiteralPath 'Preparation.json' -Raw | ConvertFrom-Json
    foreach($pin in $prep.inputs.psobject.Properties) {
        if((Get-FileHash -LiteralPath $pin.Name).Hash -cne $pin.Value){throw ('Изменён вход '+$pin.Name)}
    }
    $book='docs/regression/Diagnostic.xlsm'
    $parent=Get-Content -LiteralPath 'ParentCodeUpdate.json' -Raw | ConvertFrom-Json
    if((Test-Path -LiteralPath 'CompilePoint.log') -or (Get-FileHash -LiteralPath $book).Hash -cne $parent.WorkbookSHA256){throw 'Нарушена исходная частная книга.'}
    Assert-IsolatedOfficeReady 'CompileReadOnlyPointTrace'
    try {
        & './tools/build_workbook/Update-IsolatedOptimizerCode.ps1' -WorkbookPath $book -ModuleNames 'CCoupledArcContinuation' |
            Tee-Object -FilePath 'CompilePoint.log'
    } catch {
        [ordered]@{Message=$_.Exception.Message;ScriptStackTrace=$_.ScriptStackTrace;Position=$_.InvocationInfo.PositionMessage} |
            ConvertTo-Json -Depth 4 | Set-Content -LiteralPath 'CompilePointFailure.json' -Encoding UTF8
        throw
    } finally {
        Save-OfficeClosure 'CompilePoint.log' 'CompilePointClosure.json'
    }
} finally {Pop-Location}
