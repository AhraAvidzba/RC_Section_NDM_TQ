"""Готовит изолированный расчёт прямоугольника по квадратуре Гаусса–Лежандра."""

from argparse import ArgumentParser
from math import cos, fsum, isclose, pi
from pathlib import Path
import json


def legendre_and_derivative(order, x):
    """Вычисляет полином Лежандра и его производную по рекуррентной формуле."""
    previous, current = 1.0, x
    for degree in range(2, order + 1):
        previous, current = current, ((2 * degree - 1) * x * current - (degree - 1) * previous) / degree
    return current, order * (previous - x * current) / (1 - x * x)


def gauss_legendre(order):
    """Находит симметричные корни и проверяет веса по точным интегралам степеней."""
    nodes = [0.0] * order
    weights = [0.0] * order
    for index in range((order + 1) // 2):
        x = cos(pi * (index + 0.75) / (order + 0.5))
        for _ in range(15):
            value, derivative = legendre_and_derivative(order, x)
            update = value / derivative
            x -= update
            if abs(update) < 2e-16:
                break
        else:
            raise ArithmeticError(f"Failed to resolve root {index} for order {order}")
        _, derivative = legendre_and_derivative(order, x)
        weight = 2 / ((1 - x * x) * derivative**2)
        nodes[index], nodes[-1 - index] = -x, x
        weights[index] = weights[-1 - index] = weight
    if order % 2:
        nodes[order // 2] = 0.0
    for degree in range(2 * order):
        exact = 0.0 if degree % 2 else 2 / (degree + 1)
        check = fsum(w * x**degree for x, w in zip(nodes, weights))
        if abs(check - exact) > 2e-14:
            raise ArithmeticError(f"Monomial {degree} is not integrated exactly at order {order}")
    return nodes, weights


def rectangle_moments(nodes, weights):
    """Интегрирует геометрические характеристики без материала и сплайнов."""
    width, height = 180.0, 120.0
    half_x, half_y = width / 2, height / 2
    return {
        "area_mm2": half_x * half_y * fsum(weights) ** 2,
        "Ix_mm4": half_x * half_y**3 * fsum(weights) * fsum(w * x * x for x, w in zip(nodes, weights)),
        "Iy_mm4": half_x**3 * half_y * fsum(weights) * fsum(w * x * x for x, w in zip(nodes, weights)),
    }


def vba_float(value):
    """Записывает число Double как литерал VBA с инвариантным разделителем."""
    return format(value, ".17g").replace("e", "E") + "#"


def experiment_source(source, order, nodes, weights):
    """Заменяет только квадратуру бетона в отдельной прямоугольной модели."""
    start = source.index("    Dim mesh As New CFiberMeshBuilder")
    end = source.index("    ReDim x(1 To count)", start)
    prepare = f"""    ' -------------------- ДЛЯ ТЕСТОВ --------------------
    ' Тензорное правило Гаусса только для сплошной прямоугольной области HardShear.
    Dim order As Long: order = {order}
    Dim nodes As Variant: nodes = Array({', '.join(map(vba_float, nodes))})
    Dim weights As Variant: weights = Array({', '.join(map(vba_float, weights))})
    Dim spansX As Long: spansX = mSplineX.Count - 3
    Dim spansY As Long: spansY = mSplineY.Count - 3
    If Abs(mRegion.Area - mGeometry.Width * mGeometry.Height) > 1E-10 * mRegion.Area Then _
        GeometryError "Опытная квадратура предназначена для сплошного прямоугольника."
    Dim stepX As Double, stepY As Double
    stepX = mGeometry.Width / spansX
    stepY = mGeometry.Height / spansY
    Dim count As Long: count = order * order * spansX * spansY
"""
    source = source[:start] + prepare + source[end:]
    start = source.index("    For i = 1 To count")
    end = source.index("    If Abs(totalArea", start)
    integration = """    Dim ix As Long, iy As Long, cellX As Long, cellY As Long
    For iy = 0 To order * spansY - 1
        cellY = iy \\ order
        worldY = mGeometry.MinY + stepY * (cellY + 0.5 * (1# + nodes(iy Mod order)))
        For ix = 0 To order * spansX - 1
            cellX = ix \\ order
            worldX = mGeometry.MinX + stepX * (cellX + 0.5 * (1# + nodes(ix Mod order)))
            i = i + 1
            x(i) = worldX - mFrame.OriginX
            y(i) = worldY - mFrame.OriginY
            area(i) = 0.25 * stepX * stepY * weights(ix Mod order) * weights(iy Mod order)
            totalArea = totalArea + area(i)
            EvaluateShape x(i), y(i), localIndex, shape, dx, dy
            For k = 1 To 16
                indices(k, i) = localIndex(k)
                values(k, i) = shape(k)
                gradientX(k, i) = dx(k)
                gradientY(k, i) = dy(k)
            Next k
        Next ix
    Next iy
"""
    return source[:start] + integration + source[end:]


def main():
    """Проверяет моменты прямоугольника и сохраняет исходник частной проверки."""
    parser = ArgumentParser(description=__doc__)
    parser.add_argument("--order", type=int, required=True, choices=(5, 7, 8))
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    source_path = Path(__file__).resolve().parent.parent / "src/Solver/CCoupledSectionQuadrature.cls"
    source = source_path.read_text(encoding="utf-8-sig")
    nodes, weights = gauss_legendre(args.order)
    exact = {"area_mm2": 21600.0, "Ix_mm4": 25920000.0, "Iy_mm4": 58320000.0}
    gauss = rectangle_moments(nodes, weights)
    for name, value in exact.items():
        if not isclose(gauss[name], value, rel_tol=1e-13):
            raise ArithmeticError(f"Rectangle moment {name} is not integrated exactly")
    midpoint_nodes = [-1 + (2 * index + 1) / args.order for index in range(args.order)]
    midpoint = rectangle_moments(midpoint_nodes, [2 / args.order] * args.order)
    args.output.mkdir(parents=True, exist_ok=True)
    (args.output / source_path.name).write_text(experiment_source(source, args.order, nodes, weights), encoding="utf-8")
    (args.output / "Quadrature.json").write_text(json.dumps({
        "order_per_span": args.order, "concrete_points_one_span": args.order**2,
        "nodes": nodes, "weights": weights, "exact_rectangle": exact,
        "midpoint": midpoint, "GaussLegendre": gauss
    }, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
