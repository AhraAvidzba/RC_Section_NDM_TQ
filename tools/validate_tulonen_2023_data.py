"""Сверяет силы P1–P4 с моментами и считывает ступени с рис. 12–13."""

import argparse
from decimal import Decimal
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from statistics import median

import fitz
from PIL import Image


PDF_NAME = "Tulonen-Laaksonen-2023-Bending-Torsion-Tests.pdf"
PDF_SHA = "6797ec28088c5bd33954c7140af33201ebec32ff046d2354dd4c26bf433dd7cf"

PEAKS = (
    # Сила домкратов, эксцентриситет опор, Mx и T из табл. 2.
    ("71.2", "0.250", "28.5", "8.9"),
    ("98.0", "0.125", "39.2", "6.1"),
    ("145.2", "0.250", "58.1", "18.2"),
    ("210.2", "0.125", "84.1", "13.1"),
)


def curve_color(sample, rgb):
    """Выбирает насыщенный цвет балки вне легенды графика."""
    red, green, blue = rgb
    if sample == 0:
        return green > 150 and blue > 150 and red + 80 < min(green, blue)
    if sample == 1:
        return red > 170 and green > 100 and blue + 75 < green
    if sample == 2:
        return blue > 105 and blue > max(red, green) + 95
    return green > 78 and green > red + 55 and green > blue + 20


def moment_coordinate(image, sample, panel, moment, moment_max):
    """Оценивает абсциссу по медиане цвета и возвращает полный диапазон полосы."""
    left, right, bottom, top, x_max = panel
    row = round(bottom - (bottom - top) * moment / moment_max)
    pixels = [
        x
        for y in range(row - 2, row + 3)
        for x in range(round(left) + 1, round(right) - 1)
        if curve_color(sample, image[x, y])
    ]
    if not pixels:
        raise ValueError(f"Цвет P{sample + 1} отсутствует около строки {row}")
    return tuple((x - left) * x_max / (right - left) for x in (min(pixels), median(pixels), max(pixels)))


def verify(path):
    """Проверяет источники координат, статическое преобразование и ступени."""
    pdf = path.read_bytes()
    if sha256(pdf).hexdigest() != PDF_SHA:
        raise ValueError(f"Для считывания нужен архивный PDF с SHA {PDF_SHA}")

    # Растры извлекаются в исходном разрешении. Рамки заданы в пикселях,
    # последняя координата в панели — физическое значение правой границы.
    with fitz.open(stream=pdf, filetype="pdf") as source:
        images = source[14].get_images(full=True)
        torsion, bending = (
            Image.open(BytesIO(source.extract_image(entry[0])["image"])).convert("RGB")
            for entry in images
        )
    if torsion.size != (2084, 979) or bending.size != (1839, 871):
        raise ValueError("Растры рис. 12–13 имеют другое разрешение")

    torsion_panels = ((105.5, 981.5, 709.5, 13, .06), (1171.5, 2047.5, 709.5, 13, .06))
    bending_panels = ((93.5, 886, 631.5, 13, .09), (1041, 1833.5, 623.5, 13, .09))
    torsion_pixels, bending_pixels = torsion.load(), bending.load()
    print("beam;fraction;kappa_min;kappa_median;kappa_max;twist_min;twist_median;twist_max")
    for sample, raw in enumerate(PEAKS):
        force, eccentricity, published_bending, published_torque = map(Decimal, raw)
        peak_bending = Decimal(".4") * force
        peak_torque = Decimal(".5") * eccentricity * force
        if max(abs(peak_bending - published_bending), abs(peak_torque - published_torque)) > Decimal(".05"):
            raise ValueError(f"Моменты P{sample + 1} не совпали со схемой испытания")
        panel = sample % 2
        for fraction in (.5, .75, .9):
            kappa = moment_coordinate(bending_pixels, sample, bending_panels[panel],
                                      fraction * float(peak_bending), 90)
            twist = moment_coordinate(torsion_pixels, sample, torsion_panels[panel],
                                      fraction * float(peak_torque), 30)
            values = ";".join(f"{value:.6f}" for value in (*kappa, *twist))
            print(f"P{sample + 1};{fraction:.2f};{values}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--pdf", type=Path, default=Path("literature") / "Coupled6D" / PDF_NAME)
    verify(parser.parse_args().pdf)
