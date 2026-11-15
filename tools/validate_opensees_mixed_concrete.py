"""Сравнивает пространственный отклик материалов на одинаковой заданной истории."""

import argparse
from importlib.metadata import version
import json
from math import isclose, isfinite
from pathlib import Path

import openseespy.opensees as ops


MODULUS = 30_000.0
POISSON = 0.2
LENGTH = 100.0
TRAJECTORIES = {"Compression": -0.00018, "Shear": 0.0, "Tension": 0.00009}
PEAK_SHEAR_XZ = 0.00048
PEAK_SHEAR_YZ = 0.00024
STRESS_TOLERANCE = 2e-8  # МПа для однородности элемента и линейного эталона.
MINIMUM_COUPLING = 0.05  # МПа для выявления связи на последней точке первичного нагружения.


def schedule(steps):
    """Нагружает до 1, разгружает до 0,6 и повторно нагружает до 1."""
    valley = 3 * steps // 5
    return [0] + [index / steps for index in range(1, steps + 1)] + [
        index / steps for index in range(steps - 1, valley - 1, -1)
    ] + [index / steps for index in range(valley + 1, steps + 1)]


def solid_concrete_response(axial, factors, elastic=False):
    """Решает один куб stdBrick под однородными перемещениями с моделью ASDConcrete3D."""
    ops.wipe()
    ops.model("basic", "-ndm", 3, "-ndf", 3)
    if elastic:
        ops.nDMaterial("ElasticIsotropic", 1, MODULUS, POISSON)
    else:
        ops.nDMaterial("ASDConcrete3D", 1, MODULUS, POISSON, "-fc", 30.0, "-ft", 3.0)
    ops.timeSeries("Path", 1, "-dt", 1.0, "-values", *factors, "-useLast")
    ops.pattern("Plain", 1, 1)
    corners = ((0, 0, 0), (1, 0, 0), (1, 1, 0), (0, 1, 0), (0, 0, 1), (1, 0, 1), (1, 1, 1), (0, 1, 1))
    for tag, corner in enumerate(corners, start=1):
        x, y, z = (LENGTH * coordinate for coordinate in corner)
        ops.node(tag, x, y, z)
        displacement = (
            -POISSON * axial * x + 0.5 * PEAK_SHEAR_XZ * z,
            -POISSON * axial * y + 0.5 * PEAK_SHEAR_YZ * z,
            axial * z + 0.5 * (PEAK_SHEAR_XZ * x + PEAK_SHEAR_YZ * y),
        )
        for component, value in enumerate(displacement, start=1):
            ops.sp(tag, component, value)
    ops.element("stdBrick", 1, *range(1, 9), 1)

    # Несвязанный пружинный элемент оставляет одно уравнение в системе OpenSees.
    # Все степени свободы самого куба заданы и пружина на них не реагирует.
    for tag in (9, 10):
        ops.node(tag, -LENGTH, 0.0, 0.0)
    ops.fix(9, 1, 1, 1)
    ops.fix(10, 0, 1, 1)
    ops.uniaxialMaterial("Elastic", 2, 1.0)
    ops.element("zeroLength", 2, 9, 10, "-mat", 2, "-dir", 1)

    ops.constraints("Transformation")
    ops.numberer("Plain")
    ops.system("BandGeneral")
    ops.algorithm("Linear")
    ops.integrator("LoadControl", 1.0)
    ops.analysis("Static")
    frames = []
    max_uniformity_error = 0.0
    max_face_error = 0.0
    for factor in factors[1:]:
        if ops.analyze(1) != 0:
            raise ValueError(f"OpenSees: сбой при множителе {factor}")
        stress = ops.eleResponse(1, "material", 1, "stress")
        damage = [0.0, 0.0] if elastic else ops.eleResponse(1, "material", 1, "damage")
        strain = ops.eleResponse(1, "material", 1, "strain")
        if len(stress) != 6 or len(strain) != 6 or len(damage) != 2 or not all(map(isfinite, stress + strain + damage)):
            raise ValueError("OpenSees не вернул шесть конечных компонент и две меры повреждения")
        # Voigt в OpenSees: xx, yy, zz, xy, yz, zx. В протоколе: xx, yy, zz, xy, xz, yz.
        expected_strain = (-POISSON * axial, -POISSON * axial, axial, 0.0, PEAK_SHEAR_YZ, PEAK_SHEAR_XZ)
        if max(abs(actual - factor * expected) for actual, expected in zip(strain, expected_strain)) > 1e-12:
            raise ValueError("Куб не воспроизводит заданный тензор деформации")
        for point in range(2, 9):
            other = ops.eleResponse(1, "material", point, "stress")
            max_uniformity_error = max(max_uniformity_error, *(abs(a - b) for a, b in zip(stress, other, strict=True)))
        ops.reactions()
        for component, value in enumerate((stress[5], stress[4], stress[2]), start=1):
            face = sum(ops.nodeReaction(node, component) for node in range(5, 9)) / LENGTH**2
            max_face_error = max(max_face_error, abs(value - face))
        frames.append([factor, *stress[:4], stress[5], stress[4], *damage])

    if max(max_uniformity_error, max_face_error) > STRESS_TOLERANCE:
        raise ValueError("Куб имеет неоднородные напряжения или реакции грани не подтверждают тензор")
    return frames, max_uniformity_error, max_face_error


def elastic_error(frames, axial):
    """Сверяет первую упругую точку с независимо вычисленными E·ε и G·γ."""
    first = frames[0]
    shear_modulus = MODULUS / (2 * (1 + POISSON))
    reference = (0.0, 0.0, MODULUS * axial, 0.0, shear_modulus * PEAK_SHEAR_XZ, shear_modulus * PEAK_SHEAR_YZ)
    return max(abs(actual - first[0] * expected) for actual, expected in zip(first[1:7], reference, strict=True))


def read_native_report(path, factors):
    """Проверяет полноту заданной истории, прежде чем сравнить её с OpenSees."""
    runs = {case: [] for case in TRAJECTORIES}
    with path.open(encoding="utf-8-sig") as source:
        for record in source:
            if record.startswith("MIXED_POINT;"):
                _, case, index, *columns = record.strip().split(";")
                frame = [float(value) for value in columns]
                if len(frame) != 9 or not all(map(isfinite, frame)) or int(index) != len(runs[case]) + 1:
                    raise ValueError(f"{path}: повреждена последовательность точек {case}")
                runs[case].append(frame)
    for case, points in runs.items():
        if len(points) != len(factors) - 1 or any(
            not isclose(value, point[0], rel_tol=0.0, abs_tol=1e-12)
            for value, point in zip(factors[1:], points, strict=True)
        ):
            raise ValueError(f"{path}: OpenSees и NDM получили разные этапы траектории {case}")
    return runs


def characterize(runs, steps):
    """Сравнивает напряжения между разными осевыми деформациями и при разгрузке."""
    difference = runs["Compression"][steps - 1][5] - runs["Tension"][steps - 1][5]
    if difference < MINIMUM_COUPLING:
        raise ValueError(f"Изменение касательного напряжения при сжатии и растяжении всего {difference} МПа")
    valley = 3 * steps // 5
    return {
        "compression_minus_tension_sigma_xz_at_first_maximum_strain_MPa": difference,
        "sigma_xz_at_first_maximum_strain_MPa": {case: frames[steps - 1][5] for case, frames in runs.items()},
        "sigma_xz_on_unloading_minus_loading_at_0_6_MPa": {
            case: frames[2 * steps - valley - 1][5] - frames[valley - 1][5] for case, frames in runs.items()
        },
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--steps", type=int, choices=(40, 160), default=40)
    parser.add_argument("--native-report", type=Path, help="Протокол VBA с той же историей")
    args = parser.parse_args()
    factors = schedule(args.steps)
    runs = {}
    diagnostics = {}
    try:
        for case, axial in TRAJECTORIES.items():
            frames, uniformity, face_balance = solid_concrete_response(axial, factors)
            elastic_frames, _, _ = solid_concrete_response(axial, factors[:2], elastic=True)
            reference_error = elastic_error(elastic_frames, axial)
            if reference_error > STRESS_TOLERANCE:
                raise ValueError(f"Линейный эталон OpenSees не подтверждён, невязка {reference_error} МПа")
            diagnostics[case] = {
                "max_stress_nonuniformity_MPa": uniformity,
                "max_face_reaction_stress_error_MPa": face_balance,
                "elastic_isotropic_control_error_MPa": reference_error,
                "asd_first_step_difference_from_elastic_MPa": elastic_error(frames, axial),
            }
            runs[case] = frames
        result = {
            "openseespy": version("openseespy"),
            "engine": ops.version(),
            "initial_loading_steps": args.steps,
            "frames_per_trajectory": len(factors) - 1,
            "opensees_diagnostics": diagnostics,
            "opensees_comparison": characterize(runs, args.steps),
            "opensees_columns": ["factor", "xx", "yy", "zz", "xy", "xz", "yz", "tension_damage", "compression_damage"],
            "opensees_result": runs,
        }
        if args.native_report:
            native = read_native_report(args.native_report, factors)
            result["ndm_first_step_elastic_error_MPa"] = {
                case: elastic_error(native[case], axial) for case, axial in TRAJECTORIES.items()
            }
            if max(result["ndm_first_step_elastic_error_MPa"].values()) > STRESS_TOLERANCE:
                raise ValueError("NDM не подтвердил линейный эталон")
            result["ndm_comparison"] = characterize(native, args.steps)
            result["ndm_minus_opensees_sigma_xz_at_first_maximum_strain_MPa"] = {
                case: native[case][args.steps - 1][5] - frames[args.steps - 1][5] for case, frames in runs.items()
            }
        print(json.dumps(result, allow_nan=False, indent=2))
    finally:
        ops.wipe()


if __name__ == "__main__":
    main()
