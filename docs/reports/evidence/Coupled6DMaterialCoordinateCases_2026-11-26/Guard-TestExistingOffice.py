"""ДЛЯ ТЕСТОВ. Защищает существующие внешние Excel, сохраняя запрет живого собственного компилятора и ресурсные ограничения."""

from argparse import ArgumentParser
from hashlib import sha256
import json
from pathlib import Path


def TestGuard(root):
    """Новый Office создаётся отдельно; внешние окна и COM-процессы не становятся собственностью теста по родителю процесса."""
    name = 'Guard-IsolatedOffice.ps1'
    path = root / name
    code = path.read_text('utf-8-sig')
    start = code.index('    foreach($item in $others) {')
    end = code.index('    $unknown=@(', start)
    code = code[:start] + '''    # ДЛЯ ТЕСТОВ: все существующие экземпляры внешние для нового запуска.
    # Живой собственный компилятор исключается по точным PID и времени.
    $ownCompiler=@()
    if (Test-Path -LiteralPath 'Compile.log') {
        foreach($line in @(Get-Content -LiteralPath 'Compile.log' -Encoding utf8)) {
            if ($line -match '^OWNED_EXCEL: id=(\\d+); utcStartTicks=(\\d+)$') {
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
''' + code[end:]
    code = code.replace("Name='ProtectedUser'", "Name='ProtectedExisting'")
    path.write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep_path = root / 'Preparation.json'
    prep = json.loads(prep_path.read_text('utf-8-sig'))
    prep['inputs'][name] = sha256(path.read_bytes()).hexdigest().upper()
    (root / Path(__file__).name).write_bytes(Path(__file__).read_bytes())
    prep['inputs'][Path(__file__).name] = sha256(Path(__file__).read_bytes()).hexdigest().upper()
    prep['directed_scope']['all_preexisting_external_Excel_preserved'] = True
    prep['directed_scope']['own_live_compiler_still_blocks_launch'] = True
    prep_path.write_text(json.dumps(prep, indent=2, sort_keys=True)+'\n', encoding='utf-8')
    print('EXISTING EXTERNAL EXCEL PROTECTED; OWN IDENTITY AND MEMORY GATES RETAINED')


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--root', type=Path, required=True)
    args = parser.parse_args()
    TestGuard(args.root.resolve())
