"""Сверяет T(θ) NDM с векторными кривыми TBU3 на рис. 14b Kagermanov и Ceresa."""

import argparse
import json
import re
from hashlib import sha256
from pathlib import Path

import fitz


PDF_SHA256 = "7d9b0fd7d0bc62e5d77fa7a7727733bc109ef8c3ad33149e2f50b7f4c64f08b7"
TWISTS = (0.001, 0.002, 0.004, 0.006, 0.008, 0.010, 0.015, 0.020, 0.025, 0.028)


def figure_curves(pdf_path):
    """Читает векторы страницы 10 и калибрует оси по отметкам 0, 0,03 и 250."""
    pdf_bytes = pdf_path.read_bytes()
    if sha256(pdf_bytes).hexdigest() != PDF_SHA256:
        raise ValueError("Калибровка относится к зафиксированной версии PDF")
    curves = {}
    with fitz.open(stream=pdf_bytes, filetype="pdf") as document:
        for drawing in document[9].get_drawings():
            boundary = drawing["rect"]
            if not (323 < boundary.x0 < 324 and 299 < boundary.y0 < 306 and boundary.width > 110):
                continue
            items = drawing["items"]
            if all(item[0] == "c" for item in items):
                curves["experiment_kNm"] = [
                    cubic_point(item[1:], step / 128)
                    for item in items for step in range(129)
                ]
            elif all(item[0] == "l" for item in items):
                curves["kagermanov_ceresa_kNm"] = [items[0][1]] + [item[-1] for item in items]
    if len(curves) != 2:
        raise ValueError("На рисунке не обнаружены две кривые")
    return {
        name: [((x - 323.648) * 0.03 / (456.480 - 323.648),
                (378.778 - y) * 250 / (378.778 - 285.596)) for x, y in points]
        for name, points in curves.items()
    }


def cubic_point(points, s):
    """Вычисляет точку опубликованной кривой Безье в координатах страницы."""
    a = 1 - s
    weights = (a**3, 3 * s * a**2, 3 * s**2 * a, s**3)
    return tuple(sum(w * point[axis] for w, point in zip(weights, points)) for axis in (0, 1))


def native_curve(report_path):
    """Берёт подтверждённые точки силового этапа и следующего за ним этапа по углу."""
    points = []
    for line in report_path.read_text(encoding="utf-8-sig").splitlines():
        parts = line.split(";")
        if len(parts) == 11 and re.fullmatch(r"(?:Twist)?\d+", parts[0]) and parts[-1] == "Accepted":
            points.append((float(parts[5]), float(parts[3])))
    if not points or any(right[0] <= left[0] for left, right in zip(points, points[1:])):
        raise ValueError("Отсутствует возрастающая последовательность закручивания")
    return points


def interpolate(points, x):
    """Интерполирует на опубликованном участке без экстраполяции."""
    for (x0, y0), (x1, y1) in zip(points, points[1:]):
        if x0 <= x <= x1 and x1 > x0:
            return y0 + (y1 - y0) * (x - x0) / (x1 - x0)
    return None


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pdf", type=Path)
    parser.add_argument("--native-report", type=Path)
    args = parser.parse_args()
    curves = figure_curves(args.pdf)
    if args.native_report:
        curves["ndm_kNm"] = native_curve(args.native_report)
    table = []
    for twist in TWISTS:
        row = {"twist_rad_per_m": twist}
        row.update({name: interpolate(points, twist) for name, points in curves.items()})
        if row.get("ndm_kNm") is not None:
            row["ndm_minus_experiment_kNm"] = row["ndm_kNm"] - row["experiment_kNm"]
        table.append(row)
    print(json.dumps(table, indent=2, allow_nan=False))


if __name__ == "__main__":
    main()
