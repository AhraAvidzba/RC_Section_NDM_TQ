"""Восстанавливает усилия в сечении k опыта БВ-КВ-410(2) по статике установки."""

import argparse
from decimal import Decimal, ROUND_HALF_UP
import json
from math import cosh, fsum, hypot, pi, tanh
from pathlib import Path


D = Decimal
REFERENCE_REACTION = D("30.5")
CRACK_REACTION = D("13.3")
LEVER_K = D("0.300")
LEVER_0_6 = D("0.600")
ECCENTRICITY = D("0.410")


def section_forces(reaction):
    """Использует две равные нагрузки из рис. 2 Колчунова и Протченко."""
    force = reaction / 2
    return {
        "reference_reaction_kN": reaction,
        "moment_k_kNm": force * LEVER_K,
        "moment_at_0_6m_kNm": force * LEVER_0_6,
        "torque_kNm": force * ECCENTRICITY,
        "shear_kN": force,
    }


def elastic_face_reference(moment_kNm, torque_kNm):
    """Считает главное напряжение в середине растянутой грани свободного квадрата."""
    side = 200.0
    odd = range(1, 200, 2)
    # Из решения Прандтля для квадрата с нулевой функцией напряжений на гранях.
    torsion_constant = side**4 / 3 * (1 - 192 / pi**5 * fsum(tanh(n * pi / 2) / n**5 for n in odd))
    shear_factor = 1 - 8 / pi**2 * fsum(1 / (n**2 * cosh(n * pi / 2)) for n in odd)
    sigma = 6 * float(moment_kNm) * 1e6 / side**3
    tau = float(torque_kNm) * 1e6 * side * shear_factor / torsion_constant
    principal = (sigma + hypot(sigma, 2 * tau)) / 2
    # В середине грани на стороне растяжения вклад Qy равен нулю по симметрии
    # и условию свободной от усилий грани. Сталь и повреждение здесь отсутствуют.
    factor = 2.2 / principal
    return {
        "torsion_constant_mm4": torsion_constant,
        "bending_stress_MPa": sigma,
        "torsional_shear_MPa": tau,
        "maximum_principal_stress_at_face_MPa": principal,
        "factor_at_declared_tensile_strength": factor,
        "equivalent_reaction_kN": factor * float(REFERENCE_REACTION),
    }


def check_native_report(path):
    """Повторно проверяет шесть интегралов Excel по статике независимо от VBA."""
    records = path.read_text(encoding="utf-8-sig").splitlines()
    if "FIXTURE=TRAVUSH_SQUARE_K;M_KNM=4.575;T_KNM=6.2525;Q_KN=15.25" not in records:
        raise ValueError(f"{path}: отчёт другого расчётного примера")
    # Силы из кН в Н, моменты из кН·м в Н·мм с делением на сторону 200 мм.
    weights = (D(1000), D(5000), D(5000), D(5000), D(1000), D(1000))
    stages = []
    first_tension_bracket = None
    last_factor = D(0)
    max_residual = D(0)
    for record in records:
        if not record.startswith("Accepted;"):
            continue
        values = [D(value) for value in record.split(";")[1:]]
        if len(values) != 15 or not all(value.is_finite() for value in values):
            raise ValueError(f"{path}: неполная или некорректная строка принятого этапа")
        factor, reported_reaction = values[:2]
        reaction = REFERENCE_REACTION * factor
        if not last_factor < factor <= 1 or abs(reported_reaction - reaction) > D("1e-12"):
            raise ValueError(f"{path}: некорректный принятый этап λ={factor}")
        actual = values[2:8]
        loads = section_forces(reaction)
        expected = (D(0), loads["moment_k_kNm"], D(0), loads["torque_kNm"], D(0), loads["shear_kN"])
        residual = max(abs(got - want) * scale for got, want, scale in zip(actual, expected, weights)) / 40000
        if residual > D("0.000005"):
            raise ValueError(f"{path}: не подтверждено равновесие при λ={factor}, невязка {residual} МПа")
        if abs(residual - values[-1]) > D("1e-12"):
            raise ValueError(f"{path}: независимая невязка этапа λ={factor} не совпадает с протоколом")
        if first_tension_bracket is None and values[10] > D("0.000001"):
            first_tension_bracket = [stages[-1] if stages else D(0), reaction]
        stages.append(reaction)
        last_factor = factor
        max_residual = max(max_residual, residual)

    if not stages or f"TARGET_REACHED={last_factor == 1}" not in records:
        raise ValueError(f"{path}: отсутствуют принятые этапы или итог им противоречит")
    return {
        "file": str(path),
        "checked_prescribed_stages": len(stages),
        "max_equivalent_resultant_residual_MPa": max_residual,
        "last_checked_reaction_kN": stages[-1],
        "first_sample_above_tension_energy_threshold_reaction_bracket_kN": first_tension_bracket,
        "nominal_stage_reached": last_factor == 1,
    }


def main():
    """Сверяет статический вектор с округлёнными числами (41), (46), (48)."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--native-report", type=Path, nargs="+", help="Сверить интегралы из протоколов Excel")
    args = parser.parse_args()
    reference = section_forces(REFERENCE_REACTION)
    if tuple(
        reference[key].quantize(D("0.01"), rounding=ROUND_HALF_UP)
        for key in ("moment_k_kNm", "torque_kNm", "shear_kN")
    ) != (D("4.58"), D("6.25"), D("15.25")):
        raise ValueError("Расчётный рычаг или реакция не соответствуют опорному этапу")

    # В пространственной модели статьи три усилия получены раздельно.
    separate_model = {
        "moment_k_kNm": D("4.08"),
        "torque_kNm": D("5.96"),
        "shear_kN": D("13.32"),
    }
    factors = {key: model / reference[key] for key, model in separate_model.items()}
    result = {
        "reference_stage": reference,
        "nominal_stage_at_observed_beam_cracking": section_forces(CRACK_REACTION),
        "nominal_cracking_factor": CRACK_REACTION / REFERENCE_REACTION,
        "published_spatial_crack_model": separate_model,
        "individual_model_to_test_ratios": factors,
        "elastic_plain_square_face_reference": elastic_face_reference(reference["moment_k_kNm"], reference["torque_kNm"]),
    }
    if args.native_report:
        result["native_resultant_checks"] = [check_native_report(path) for path in args.native_report]
    print(json.dumps(result, default=float, indent=2, allow_nan=False))


if __name__ == "__main__":
    main()
