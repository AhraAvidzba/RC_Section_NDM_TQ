"""Сверяет поля НДМ с векторными кривыми TBU3 на рис. 14–15 Kagermanov и Ceresa."""

import argparse
import json
import math
import re
from hashlib import sha256
from pathlib import Path

import fitz


PDF_SHA256 = "7d9b0fd7d0bc62e5d77fa7a7727733bc109ef8c3ad33149e2f50b7f4c64f08b7"
RASP_PDF_SHA256 = "93612557981758e2fc5c0396df3cf1820da0c2d23e06a8cd41c91b9964c9b23b"
TWISTS = (0.001, 0.002, 0.004, 0.006, 0.008, 0.010, 0.015, 0.020, 0.025, 0.028)
CURVATURES = (0.0005, 0.001, 0.002, 0.003, 0.004, 0.005, 0.006, 0.007, 0.0075, 0.008, 0.009, 0.0095)
LOAD_STAGES = (4, 8, 12, 15, 16)
HOOP_TORQUES = (15, 30, 50, 80, 100, 120, 125, 145, 165, 180, 195)
INITIAL_MOMENT = 29.6
BENDING_STEP = 14.87
TORSION_STEP = 10.35
# Полосы пикселей из рис. 9a. Первая относится к опыту, вторая к RASP.
RASP_READINGS = (
    (0.004, (181, 187), (188, 193)),
    (0.006, (169, 175), (179, 184)),
    (0.008, (157, 163), (168, 174)),
    (0.010, (148, 153), (161, 165)),
    (0.012, (138, 143), (150, 155)),
    (0.015, (123, 128), (138, 143)),
    (0.018, (110, 115), (127, 132)),
    (0.020, (102, 107), (120, 124)),
    (0.022, (94, 99), (112, 118)),
    (0.024, (88, 93), (106, 111)),
    (0.026, (83, 88), None),
    (0.028, None, (93, 98)),
    (0.030, None, (89, 94)),
)


def figure_curves(pdf_path, by_curvature=False):
    """Читает векторные линии двух кривых и калибрует оси по делениям графика."""
    pdf_bytes = pdf_path.read_bytes()
    if sha256(pdf_bytes).hexdigest() != PDF_SHA256:
        raise ValueError("Калибровка относится к зафиксированной версии PDF")
    curves = {}
    with fitz.open(stream=pdf_bytes, filetype="pdf") as document:
        for drawing in document[9].get_drawings():
            boundary = drawing["rect"]
            if by_curvature:
                if not (149 < boundary.x0 < 151 and 292 < boundary.y0 < 297 and boundary.width > 65):
                    continue
            elif not (323 < boundary.x0 < 324 and 299 < boundary.y0 < 306 and boundary.width > 110):
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
    if by_curvature:
        # По X: отметки 0 и 0,01 рад/м. По Y: внутренние деления 50 и 250 кН·м.
        return {
            name: [((x - 150.508) * 0.01 / (240.115 - 150.508),
                    50 + (364.634 - y) * 200 / (364.634 - 310.938)) for x, y in points]
            for name, points in curves.items()
        }
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


def hoop_figure_curves(pdf_path):
    """Калибрует значения деформации при росте T по осевым линиям рис. 15b."""
    source = pdf_path.read_bytes()
    if sha256(source).hexdigest() != PDF_SHA256:
        raise ValueError("Калибровка хомута относится к зафиксированной версии PDF")
    curves = {}
    with fitz.open(stream=source, filetype="pdf") as document:
        for drawing in document[9].get_drawings():
            bounds = drawing["rect"]
            if not (341 < bounds.x0 < 344 and 516 < bounds.y0 < 524 and bounds.width > 100):
                continue
            segments = drawing["items"]
            if any(segment[0] == "c" for segment in segments):
                name = "experiment"
                points = [segments[0][1]]
                for segment in segments:
                    if segment[0] == "c":
                        points.extend(cubic_point(segment[1:], step / 128) for step in range(1, 129))
                    elif segment[0] == "l":
                        points.append(segment[-1])
                    else:
                        raise ValueError("Неподдерживаемая форма отрезка рис. 15b")
            elif all(segment[0] == "l" for segment in segments):
                name = "kagermanov_ceresa"
                points = [segments[0][1]] + [segment[-1] for segment in segments]
            else:
                raise ValueError("Не распознана линия рис. 15b")
            if name in curves:
                raise ValueError(f"Найдена лишняя линия {name} на рис. 15b")
            # По X отметки −0,0002 и 0,0018, по Y отметки 0 и 250 кН·м.
            curves[name] = [((599.495 - y) * 250 / (599.495 - 497.057),
                             -0.0002 + (x - 328.452) * 0.002 / (478.303 - 328.452)) for x, y in points]
    if set(curves) != {"experiment", "kagermanov_ceresa"}:
        raise ValueError("Для хомута нужны обе линии рис. 15b")
    return curves


def native_hoop_curves(report_path):
    """Проверяет обе координаты измерения, путь нагрузки и соответствие принятому НДС."""
    rows, accepted = [], []
    gauge_axes = None
    passed = None
    for line in report_path.read_text(encoding="utf-8-sig").splitlines():
        parts = line.split(";")
        if parts[0] == "TBU3_GAUGE_X_MM" and len(parts) == 3 and gauge_axes is None:
            gauge_axes = tuple(map(float, parts[1:]))
        elif parts[0] == "TBU3_GAUGE" and len(parts) == 6:
            rows.append(tuple(map(float, parts[1:])))
        elif len(parts) == 11 and re.fullmatch(r"(?:Twist)?\d+", parts[0]) and parts[-1] == "Accepted":
            accepted.append(tuple(map(float, parts[1:4])))
        elif match := re.fullmatch(r"TOTAL_TBU3_FORCE_EQUILIBRIUM: passed=(\d+); failed=0", line):
            passed = int(match.group(1))
    if (gauge_axes is None or any(not math.isclose(actual, expected, abs_tol=1e-8, rel_tol=0.)
                                 for actual, expected in zip(gauge_axes, (-222., 222.), strict=True))
            or passed is None or passed < len(rows) or len(rows) < 2 or len(rows) != len(accepted)):
        raise ValueError("Протокол не содержит полного ряда двух измерений из принятых полей TBU3")
    for (stage, moment, torque, left, right), expected in zip(rows, accepted, strict=True):
        if not all(map(math.isfinite, (stage, moment, torque, left, right))) or (stage, moment, torque) != expected:
            raise ValueError("Деформации хомута относятся к отличной от принятой точке траектории")
        if max(abs(moment - (INITIAL_MOMENT + BENDING_STEP * stage)), abs(torque - TORSION_STEP * stage)) > 0.0002:
            raise ValueError("Силовые компоненты не совпадают с заданной траекторией TBU3")
    if abs(rows[0][0]) > 1e-8 or abs(rows[0][2]) > 0.0002:
        raise ValueError("Отсутствует принятое состояние одного предварительного момента")
    return {"left": [(torque, left) for _, _, torque, left, _ in rows],
            "right": [(torque, right) for _, _, torque, _, right in rows]}


def compare_hoop_gauge(pdf_path, report_path):
    """Сопоставляет деформации при общем T с предварительным M и без экстраполяции."""
    digitized = hoop_figure_curves(pdf_path)
    native = native_hoop_curves(report_path) if report_path else {}
    table = []
    for torque in HOOP_TORQUES:
        strains = {name: interpolate(points, torque) for name, points in digitized.items()}
        for side, points in native.items():
            value = interpolate(points, torque)
            strains[f"ndm_{side}"] = value
            strains[f"ndm_{side}_from_preload"] = value - points[0][1] if value is not None else None
        table.append({"T_kNm": torque, "strain": strains})
    return {"source_sha256": PDF_SHA256, "figure": "15b", "reading_resolution_microstrain": 5,
            "native_max_T_kNm": max((torque for torque, _ in native.get("left", [])), default=None),
            "native_preload_strain_gauges": {side: points[0][1] for side, points in native.items()}, "rows": table}


def native_curve(report_path, by_curvature=False):
    """Берёт подтверждённые точки до и после смены управления нагрузкой."""
    points = []
    abscissa, ordinate = (4, 2) if by_curvature else (5, 3)
    for line in report_path.read_text(encoding="utf-8-sig").splitlines():
        parts = line.split(";")
        if len(parts) == 11 and re.fullmatch(r"(?:Twist)?\d+", parts[0]) and parts[-1] == "Accepted":
            points.append((float(parts[abscissa]), float(parts[ordinate])))
    if not points or any(right[0] <= left[0] for left, right in zip(points, points[1:])):
        raise ValueError("Отсутствует возрастающая последовательность кривизны или закручивания")
    return points


def interpolate(points, x):
    """Интерполирует на опубликованном участке без экстраполяции."""
    for (x0, y0), (x1, y1) in zip(points, points[1:]):
        if x0 <= x <= x1 and x1 > x0:
            return y0 + (y1 - y0) * (x - x0) / (x1 - x0)
    return None


def load_curves(pdf_path, report_path, by_curvature):
    """Читает две опубликованные кривые и сохранённый ряд NDM при его наличии."""
    curves = figure_curves(pdf_path, by_curvature)
    if report_path:
        curves["ndm_kNm"] = native_curve(report_path, by_curvature)
    return curves


def compare_load_stages(pdf_path, report_path):
    """Сравнивает деформации при той же номинальной паре M, T.

    Для каждого момента берётся первое пересечение при его увеличении.
    Из двух опубликованных графиков деформации читаются независимо.
    """
    bending = load_curves(pdf_path, report_path, True)
    torsion = load_curves(pdf_path, report_path, False)
    if report_path:
        for (_, moment), (_, torque) in zip(bending["ndm_kNm"], torsion["ndm_kNm"], strict=True):
            expected_moment = INITIAL_MOMENT + BENDING_STEP * torque / TORSION_STEP
            if abs(moment - expected_moment) > 0.0002:
                raise ValueError("Сохранённый ряд NDM относится к другой траектории M и T")
    table = []
    for stage in LOAD_STAGES:
        moment = round(INITIAL_MOMENT + BENDING_STEP * stage, 2)
        torque = round(TORSION_STEP * stage, 2)
        curves = {}
        for name in bending:
            curves[name.removesuffix("_kNm")] = {
                "curvature_per_m": interpolate([(y, x) for x, y in bending[name]], moment),
                "twist_rad_per_m": interpolate([(y, x) for x, y in torsion[name]], torque),
            }
        table.append({"nominal_stage": stage, "M_kNm": moment, "T_kNm": torque, "curves": curves})
    return table


def rasp_curves(pdf_path):
    """Извлекает среднюю координату линии из размеченной полосы скана рис. 9a."""
    pdf = pdf_path.read_bytes()
    if sha256(pdf).hexdigest() != RASP_PDF_SHA256:
        raise ValueError("Разметка относится к зафиксированной копии рисунка 9a")
    with fitz.open(stream=pdf, filetype="pdf") as document:
        raster = document[6].get_pixmap(matrix=fitz.Matrix(1.6, 1.6), colorspace=fitz.csGRAY)
    if (raster.width, raster.height) != (914, 1286):
        raise ValueError("Размер растра отличается от калибровочного")

    # Положения делений 0 и 0,04 рад/м по X, 0 и 200 кН·м по Y.
    x_origin, x_at_004 = 176.5, 392.5
    y_origin, y_at_200 = 222.9, 55.7
    labels = ("experiment_aci", "rasp")
    curves = {label: [] for label in labels}
    pixels = raster.samples
    for nominal_twist, *bands in RASP_READINGS:
        x = round(x_origin + (x_at_004 - x_origin) * nominal_twist / 0.04)
        twist = (x - x_origin) * 0.04 / (x_at_004 - x_origin)
        for label, band in zip(labels, bands):
            if band is None:
                continue
            rows = [y for y in range(band[0], band[1] + 1) if pixels[y * raster.stride + x] < 112]
            if not rows:
                raise ValueError(f"В полосе {label} при x={x} линия не обнаружена")
            torque = (y_origin - sum(rows) / len(rows)) * 200 / (y_origin - y_at_200)
            curves[label].append((twist, torque))
    return curves


def compare_rasp(rasp_path, pdf_path, report_path):
    """Сопоставляет оба источника и NDM по углу и номинальным усилиям."""
    curves = load_curves(pdf_path, report_path, False)
    curves = {name.removesuffix("_kNm"): points for name, points in curves.items()}
    curves["experiment_kagermanov_figure"] = curves.pop("experiment")
    curves.update(rasp_curves(rasp_path))
    fixed_twist = [
        {"twist_per_m": twist, "torques_kNm": {name: interpolate(points, twist) for name, points in curves.items()}}
        for twist, _ in curves["experiment_aci"]
    ]
    load_stages = []
    for stage in (8, 12, 15):
        torque = round(TORSION_STEP * stage, 2)
        load_stages.append({
            "stage": stage, "moment_kNm": round(INITIAL_MOMENT + BENDING_STEP * stage, 2), "torque_kNm": torque,
            "twists_per_m": {name: interpolate([(y, x) for x, y in points], torque) for name, points in curves.items()}
        })
    return {"source_sha256": RASP_PDF_SHA256, "aci_figure_reading_uncertainty_kNm": 4,
            "fixed_twist_grid": fixed_twist, "nominal_load_stages": load_stages}


def main():
    """Проверяет выбранный набор кривых и печатает воспроизводимую таблицу."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pdf", type=Path)
    parser.add_argument("--native-report", type=Path)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--curvature", action="store_true", help="рисунок 14a, момент от кривизны")
    mode.add_argument("--load-path", action="store_true", help="деформации при общей номинальной паре M, T")
    mode.add_argument("--hoop-gauge", action="store_true", help="рисунок 15b, деформация хомута на середине стенки")
    mode.add_argument("--rasp-pdf", type=Path, help="добавить опыт и RASP из рисунка 9a")
    args = parser.parse_args()
    if args.hoop_gauge:
        print(json.dumps(compare_hoop_gauge(args.pdf, args.native_report), indent=2, allow_nan=False))
        return
    if args.rasp_pdf:
        print(json.dumps(compare_rasp(args.rasp_pdf, args.pdf, args.native_report), indent=2, allow_nan=False))
        return
    if args.load_path:
        print(json.dumps(compare_load_stages(args.pdf, args.native_report), indent=2, allow_nan=False))
        return
    curves = load_curves(args.pdf, args.native_report, args.curvature)
    coordinates = CURVATURES if args.curvature else TWISTS
    coordinate_name = "curvature_per_m" if args.curvature else "twist_rad_per_m"
    table = []
    for coordinate in coordinates:
        row = {coordinate_name: coordinate}
        row.update({name: interpolate(points, coordinate) for name, points in curves.items()})
        if row.get("ndm_kNm") is not None:
            row["ndm_minus_experiment_kNm"] = row["ndm_kNm"] - row["experiment_kNm"]
            if args.curvature and row["kagermanov_ceresa_kNm"] is not None:
                row["ndm_minus_kagermanov_ceresa_kNm"] = row["ndm_kNm"] - row["kagermanov_ceresa_kNm"]
        table.append(row)
    print(json.dumps(table, indent=2, allow_nan=False))


if __name__ == "__main__":
    main()
