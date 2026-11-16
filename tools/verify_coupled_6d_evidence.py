"""Воспроизводит зафиксированные численные проверки связанного расчёта."""

import argparse
from hashlib import sha256
from importlib.metadata import PackageNotFoundError, version
from io import BytesIO
import json
import os
from pathlib import Path, PurePosixPath
import subprocess
import sys
from tempfile import TemporaryDirectory
from zipfile import BadZipFile, ZipFile


MANIFEST_PATH = Path("docs/reports/Coupled6DEvidence.json")
ARCHIVE_DIRECTORY = Path("docs/regression/Coupled6D")


def require_digest(path, expected):
    """Проверяет точные байты входа или результата и возвращает их."""
    data = path.read_bytes()
    actual = sha256(data).hexdigest().upper()
    if actual != expected:
        raise ValueError(f"{path.name}: SHA-256 {actual}, ожидался {expected}")
    return data


def extract_archive(source, fingerprint, destination):
    """Распаковывает проверенный ZIP в отдельный чистый каталог."""
    with ZipFile(BytesIO(require_digest(source, fingerprint))) as archive:
        member_names = set()
        for entry in archive.infolist():
            name = PurePosixPath(entry.filename)
            if (
                not name.parts
                or name.is_absolute()
                or ".." in name.parts
                or ":" in entry.filename
                or "\\" in entry.filename
                or name.as_posix().casefold() in member_names
            ):
                raise ValueError(f"{source.name}: недопустимая запись {entry.filename!r}")
            member_names.add(name.as_posix().casefold())
        archive.extractall(destination)


def replay_case(case, locations):
    """Запускает один архивный проверяющий сценарий и сверяет его результат."""
    archive = Path(locations[case["archive"]])
    checks = case.get("outputs", {})
    for filename in checks:
        output = (archive / filename).resolve()
        if not output.is_relative_to(archive.resolve()):
            raise ValueError(f"{case['id']}: путь результата за каталогом распаковки")
        output.unlink(missing_ok=True)  # Удаляется только копия ожидаемого результата в распаковке.
    command = [sys.executable, "-B", "-X", "utf8", case["script"]]
    command.extend(argument.format_map(locations) for argument in case.get("arguments", []))
    environment = os.environ.copy()
    environment.pop("PYTHONOPTIMIZE", None)  # Архивные сценарии проверяют результаты через assert.
    result = subprocess.run(
        command, cwd=archive, env=environment, capture_output=True,
        encoding="utf-8", timeout=90, check=False,
    )
    if result.returncode or case.get("stdout_contains", "") not in result.stdout:
        detail = result.stderr.strip() or result.stdout.strip()
        raise ValueError(f"{case['id']}: проверка не пройдена. {detail}")
    for filename, fingerprint in checks.items():
        require_digest(archive / filename, fingerprint)
    print(f"PASS {case['id']}", flush=True)
    return {
        "id": case["id"],
        "journal_section": case["journal_section"],
        "archive": case["archive"],
        "verified_outputs_sha256": checks,
        "verified_stdout": case.get("stdout_contains"),
    }


def replay(repo):
    """Проверяет исходную книгу и воспроизводит все перечисленные проверки."""
    manifest_path = repo / MANIFEST_PATH
    content = manifest_path.read_bytes()
    manifest = json.loads(content)
    for file, fingerprint in manifest["installed_inputs"].items():
        require_digest(repo / file, fingerprint)
    archive_directory = repo / ARCHIVE_DIRECTORY
    libraries = {name: version(name) for name in manifest["python_packages"]}
    with TemporaryDirectory(prefix="ndm-coupled-evidence-") as scratch:
        locations = {"repo": str(repo), "evidence": str(archive_directory)}
        for name, fingerprint in manifest["archives"].items():
            if name != Path(name).name:
                raise ValueError(f"Имя архива должно быть простым: {name}")
            destination = Path(scratch) / name
            extract_archive(archive_directory / f"{name}.zip", fingerprint, destination)
            locations[name] = str(destination)
        cases = [replay_case(case, locations) for case in manifest["cases"]]
    return {
        "manifest_sha256": sha256(content).hexdigest().upper(),
        "source_checkpoint": manifest["source_checkpoint"],
        "installed_inputs": manifest["installed_inputs"],
        "python": sys.version.split()[0],
        "python_packages": libraries,
        "mode": "offline_replay_of_archived_evidence",
        "passed": len(cases),
        "cases": cases,
    }


def main():
    """Разбирает аргументы и сохраняет отчёт в явно указанном новом файле."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    try:
        if args.output.exists():
            raise ValueError(f"Укажите новый выходной файл вместо {args.output}")
        result = replay(args.repo.resolve())
        with args.output.open("x", encoding="utf-8", newline="\n") as report:
            json.dump(result, report, indent=2, ensure_ascii=False)
            report.write("\n")
    except (OSError, ValueError, PackageNotFoundError, BadZipFile, subprocess.TimeoutExpired) as exc:
        parser.exit(1, f"Ошибка проверки: {exc}\n")
    print(f"PASS {result['passed']} проверок. Отчёт {args.output}")


if __name__ == "__main__":
    main()
