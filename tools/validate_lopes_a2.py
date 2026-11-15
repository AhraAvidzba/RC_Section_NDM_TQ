"""Независимая упругая проверка номинального сечения A2 на чистое кручение."""

import json
from importlib.metadata import version

from sectionproperties.analysis import Section
from sectionproperties.pre import Geometry
from shapely.geometry import Polygon


OUTER = [(-300, -300), (300, -300), (300, 300), (-300, 300)]
OPENING = [
    (-150, -200), (150, -200), (200, -150), (200, 150),
    (150, 200), (-150, 200), (-200, 150), (-200, -150),
]
AREA_MM2 = 205_000
POISSON = 0.20
MEASURED_RIGIDITY_KNM2 = 97_950


def finite_element_torsion(maximum_triangle_area):
    """Считает J однородного сечения. Измеренную жёсткость использует после решения."""
    geometry = Geometry(Polygon(OUTER, [OPENING]))
    section = Section(geometry.create_mesh(mesh_sizes=maximum_triangle_area))
    section.calculate_geometric_properties()
    section.calculate_warping_properties()
    area = float(section.get_area())
    if abs(area - AREA_MM2) > AREA_MM2 * 1e-10:
        raise ValueError("Площадь сетки не совпадает с площадью принятой геометрии")
    torsion_constant = float(section.get_j())
    return {
        "max_triangle_area_mm2": maximum_triangle_area,
        "elements": len(section.elements),
        "area_mm2": area,
        "j_mm4": torsion_constant,
        "e_implied_mpa": 2 * (1 + POISSON) * MEASURED_RIGIDITY_KNM2 * 1e9 / torsion_constant,
    }


if __name__ == "__main__":
    print(json.dumps({"sectionproperties": version("sectionproperties"),
                      "meshes": [finite_element_torsion(area) for area in (800, 200, 50)]},
                     indent=2, allow_nan=False))
