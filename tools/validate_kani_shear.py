"""Сопоставляет чистый сдвиг NDM с опытом Kani и двумя расчётами из рис. 5."""

import argparse
import json
import re
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from statistics import median

import fitz
from PIL import Image


PDF_SHA256 = "0fc6dde8b392b3444fadb6a1d770348df4b5861dbe8e4b766fbd4782dd06fd83"
SHEAR_STRAINS = (0.0003, 0.0005, 0.001, 0.002, 0.003, 0.004, 0.005, 0.006, 0.0068)


def figure_curves(pdf_path):
    """Читает исходный растр из PDF, исключая оси, сетку и легенду."""
    data = pdf_path.read_bytes()
    if sha256(data).hexdigest() != PDF_SHA256:
        raise ValueError("Координаты рисунка относятся к зафиксированной версии PDF")
    with fitz.open(stream=data, filetype="pdf") as document:
        raster = document.extract_image(56)["image"]
    image = Image.open(BytesIO(raster)).convert("RGB")
    if image.size != (560, 420):
        raise ValueError("Размер растра рисунка 5 отличается от ожидаемого")
    pixels = image.load()
    curves = {"experiment_kN": [], "spline_poliotti_kN": [], "tinsa_kN": []}
    for x in range(75, 505):
        rows = {name: [] for name in curves}
        # При x >= 271 кривые идут выше легенды (верхняя граница y = 206).
        for y in range(32, 205 if x >= 271 else 352):
            red, green, blue = pixels[x, y]
            if max(red, green, blue) < 60:
                rows["experiment_kN"].append(y)
            elif red > 180 and max(green, blue) < 90:
                rows["spline_poliotti_kN"].append(y)
            elif blue > 160 and max(red, green) < 90:
                rows["tinsa_kN"].append(y)
        for name, ordinates in rows.items():
            if ordinates:
                curves[name].append(((x - 72) * 0.007 / 435, (357 - median(ordinates)) * 800 / 327))
    return curves


def native_curve(report_path):
    """Читает только подтверждённые точки на возрастающей координате сдвига."""
    points = [(0.0, 0.0)]
    for line in report_path.read_text(encoding="utf-8-sig").splitlines():
        fields = line.split(";")
        if len(fields) == 7 and re.fullmatch(r"Kani\d+", fields[0]):
            points.append((float(fields[1]), float(fields[2])))
    if len(points) < 2 or any(right[0] <= left[0] for left, right in zip(points, points[1:])):
        raise ValueError("В отчёте отсутствуют принятые шаги или нарушен порядок сдвиговых деформаций")
    return points


def interpolate(series, x):
    """Интерполирует внутри ряда. Граница сдвига имеет допуск округления VBA."""
    if abs(x - series[-1][0]) < 1e-10:
        return series[-1][1]
    for (x0, y0), (x1, y1) in zip(series, series[1:]):
        if x0 <= x <= x1:
            return y0 + (y1 - y0) * (x - x0) / (x1 - x0)
    return None


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pdf", type=Path)
    parser.add_argument("--native-report", type=Path)
    args = parser.parse_args()
    curves = figure_curves(args.pdf)
    if args.native_report:
        curves["ndm_kN"] = native_curve(args.native_report)
    table = [{"gamma": gamma, **{name: interpolate(series, gamma) for name, series in curves.items()}}
             for gamma in SHEAR_STRAINS]
    print(json.dumps(table, indent=2, allow_nan=False))


if __name__ == "__main__":
    main()
