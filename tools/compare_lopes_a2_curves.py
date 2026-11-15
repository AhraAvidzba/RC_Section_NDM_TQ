"""Сопоставляет принятые точки NDM с опубликованными расчётами A2 на рис. 8 и 10."""

import argparse
import json
import re
from hashlib import sha256
from math import degrees, radians
from pathlib import Path
from statistics import median

import fitz


PDF_SHA256 = "83bf3db6babe0d01009d1a098a0fda8bba40cb3af71939e929e55128261c8564"
TWIST_DEGREES = (0.025, 0.05, 0.064, 0.10, 0.125, 0.20, 0.50, 1.08, 1.17, 1.66)


def curve_values(drawing, twists, axes):
    """Считывает ординату в точке пересечения кривой с заданной абсциссой."""
    x_origin, x_second_tick, second_tick_deg, y_origin, y_maximum, moment_maximum = axes
    segments = []
    for kind, *points in drawing["items"]:
        if kind == "l":
            segments.append(points)
        elif kind == "c":
            samples = [cubic_point(points, i / 12) for i in range(13)]
            segments.extend(zip(samples, samples[1:]))
        else:
            raise ValueError("Неожиданное представление кривой")

    result = []
    scale = moment_maximum / (y_origin - y_maximum)
    filled_outline = drawing["type"] == "f"
    for angle in twists:
        x = x_origin + angle * (x_second_tick - x_origin) / second_tick_deg
        intersect = [y0 + (y1 - y0) * (x - x0) / (x1 - x0)
                     for (x0, y0), (x1, y1) in segments if x0 != x1 and min(x0, x1) <= x < max(x0, x1)]
        # У залитой линии два края. Другие кратности оставляют ординату неоднозначной.
        if not intersect or (filled_outline and len(intersect) != 2):
            result.append(None)
        elif not filled_outline and (max(intersect) - min(intersect)) * scale > 3:
            result.append(None)
        else:
            result.append((y_origin - median(intersect)) * scale)
    return result


def cubic_point(points, s):
    """Вычисляет координаты точки кубической кривой PDF."""
    a = 1 - s
    weights = (a**3, 3 * a**2 * s, 3 * a * s**2, s**3)
    return tuple(sum(weight * p[axis] for weight, p in zip(weights, points)) for axis in (0, 1))


def native_points(report):
    """Извлекает силу преднагружения и принятые точки со строго возрастающим углом."""
    axial_force = 0.0
    accepted = [(0.0, 0.0)]
    for line in report.read_text(encoding="utf-8-sig").splitlines():
        if line.startswith("A2_PRELOAD_N_KN:"):
            axial_force = float(line.split(":", 1)[1])
        row = line.split(";")
        if len(row) == 8 and re.fullmatch(r"A2Step\d+", row[0]):
            accepted.append((degrees(float(row[1])), float(row[2])))
    if len(accepted) < 2 or any(b[0] <= a[0] for a, b in zip(accepted, accepted[1:])):
        raise ValueError("Отчёт должен содержать возрастающий ряд принятых закручиваний")
    return axial_force, accepted


def interpolate(series, x):
    """Берёт точку только внутри принятого ряда."""
    for (x0, y0), (x1, y1) in zip(series, series[1:]):
        if x0 <= x <= x1:
            return y0 + (y1 - y0) * (x - x0) / (x1 - x0)
    return None


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pdf", type=Path)
    parser.add_argument("native_report", type=Path)
    args = parser.parse_args()
    contents = args.pdf.read_bytes()
    if sha256(contents).hexdigest() != PDF_SHA256:
        raise ValueError("Координаты рисунков относятся к зафиксированной версии PDF")
    force, accepted = native_points(args.native_report)
    if abs(force) < 1e-4:
        fem_path, truss_path = 57, 11
    elif abs(force + 1279.1) < 1e-4:
        fem_path, truss_path = 58, 13
    else:
        raise ValueError("В сравнении предусмотрены N = 0 и N = −1279,1 кН")
    with fitz.open(stream=contents, filetype="pdf") as document:
        fem = curve_values(document[13].get_drawings()[fem_path], TWIST_DEGREES,
                           (180.727, 336.787, 4, 693.079, 542.539, 350))
        truss = curve_values(document[12].get_drawings()[truss_path], TWIST_DEGREES,
                             (169.403, 379.103, 3.5, 579.573, 422.853, 450))
    result = [{"twist_deg_per_m": angle, "twist_rad_per_m": radians(angle),
               "ndm_kNm": interpolate(accepted, angle), "fem_kNm": fe, "truss_kNm": tr}
              for angle, fe, tr in zip(TWIST_DEGREES, fem, truss)]
    print(json.dumps({"normal_kN": force, "values": result}, indent=2, allow_nan=False))


if __name__ == "__main__":
    main()
