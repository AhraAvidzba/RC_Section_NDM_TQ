"""Сверяет пробные и принятые состояния симметричной арматуры VBA с Steel01."""

import argparse
import json
from math import isfinite
from pathlib import Path

import openseespy.opensees as ops


GRADES = {"L500": (200000.0, 500.0, 2000.0), "H240": (200000.0, 240.0, 2000.0)}
STRESS_TOLERANCE_MPA = 2e-10
TANGENT_TOLERANCE_MPA = 1e-8


def read_report(path):
    """Проверяет журнальные параметры, порядок строк и успешное завершение Excel."""
    lines = path.read_text(encoding="utf-8-sig").splitlines()
    required = ("TOTAL_COUPLED_STEEL_OPENSEES: passed=2; failed=0", "SOURCE_UNCHANGED: True")
    if any(tag not in lines for tag in required):
        raise ValueError("Протокол Excel не содержит успешного завершения контрольного сценария")
    laws = {}
    rows = {grade: [] for grade in GRADES}
    for line in lines:
        tag, *fields = line.split("|")
        if tag == "STEEL_LAW":
            grade, *values = fields
            if grade in laws or tuple(map(float, values)) != GRADES.get(grade):
                raise ValueError(f"Неверный или повторный закон стали {grade}")
            laws[grade] = GRADES[grade]
        if tag == "STEEL_STEP":
            grade, number, action, *values = fields
            samples = rows[grade]
            if int(number) != len(samples) or action not in ("Commit", "Probe") or len(values) != 5:
                raise ValueError(f"Неверная строка {grade} {number}")
            state = tuple(map(float, values))
            if not all(map(isfinite, state)):
                raise ValueError(f"Неконечное состояние {grade} {number}")
            samples.append((action == "Commit", *state))
    if laws != GRADES or any(len(samples) != 23 for samples in rows.values()):
        raise ValueError("Ожидаются две полные истории VBA")
    return rows


def reference(history, strain, elastic, strength, hardening):
    """Заново воспроизводит только принятый префикс, затем оценивает одну деформацию."""
    ops.wipe()
    ops.uniaxialMaterial("Steel01", 1, strength, elastic, hardening / elastic)
    ops.testUniaxialMaterial(1)
    for accepted in history:
        ops.setStrain(accepted)
    ops.setStrain(strain)
    return ops.getStress(), ops.getTangent()


def compare(grade, samples):
    """Сверяет весь путь OpenSees и отдельно замыкание закона и двух откатов VBA."""
    elastic, strength, hardening = GRADES[grade]
    kinematic_modulus = hardening / (1 - hardening / elastic)
    history = []
    counts = {"elastic": 0, "positive_yield": 0, "negative_yield": 0, "probe": 0}
    stress_error = tangent_error = closure_error = 0.0
    for accepted, strain, stress, tangent, plastic, backstress in samples:
        expected_stress, expected_tangent = reference(history, strain, elastic, strength, hardening)
        stress_error = max(stress_error, abs(expected_stress - stress))
        tangent_error = max(tangent_error, abs(expected_tangent - tangent))
        closure_error = max(closure_error, abs(stress - elastic * (strain - plastic)),
                            abs(backstress - kinematic_modulus * plastic))
        if tangent == elastic:
            counts["elastic"] += 1
        else:
            counts["positive_yield" if stress > backstress else "negative_yield"] += 1
            closure_error = max(closure_error, abs(abs(stress - backstress) - strength))
        if accepted:
            history.append(strain)
        else:
            counts["probe"] += 1
    if not all(counts.values()) or counts["probe"] != 7:
        raise ValueError(f"Неполное покрытие ветвей {grade}")
    if samples[5] != samples[7] or samples[13] != samples[15]:
        raise ValueError(f"Отклонённая проба изменила состояние {grade}")
    if max(stress_error, closure_error) > STRESS_TOLERANCE_MPA or tangent_error > TANGENT_TOLERANCE_MPA:
        raise ValueError(f"{grade}: ошибки напряжения, касательной и замыкания {stress_error, tangent_error, closure_error}")
    return {"maximum_stress_error_mpa": stress_error, "maximum_tangent_error_mpa": tangent_error,
            "maximum_constitutive_error_mpa": closure_error, "counts": counts, "accepted": len(history)}


def main():
    """Сохраняет количественное сравнение без изменения протокола VBA."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("report", type=Path)
    parser.add_argument("--out", type=Path, required=True)
    args = parser.parse_args()
    rows = read_report(args.report)
    try:
        result = {"opensees_version": ops.version(), "grades": {grade: compare(grade, states)
                                                              for grade, states in rows.items()}}
    finally:
        ops.wipe()
    output = json.dumps(result, ensure_ascii=False, indent=2, allow_nan=False) + "\n"
    with args.out.open("x", encoding="utf-8") as stream:
        stream.write(output)
    print(output, end="")


if __name__ == "__main__":
    main()
