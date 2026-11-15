"""Независимая проверка отклонённых проб бетонной точки 7 балки A2."""

import json
from math import acos, cos, exp, expm1, hypot, log, pi, radians, sin, sqrt, tan


FC = 47.3
FT = 3.5
ALPHA = 0.12
MODULUS = 36_100.0
POISSON = 0.2
SHEAR2 = MODULUS / (1 + POISSON)
BULK9 = 3 * MODULUS / (1 - 2 * POISSON)
ECCENTRIC_OFFSET = 0.1 * FT
PHI = radians(65)
PSI = radians(32)
COMPRESSION_PLASTIC = 1.33878957702629e-6
EARLY_TENSION_PLASTIC = 1.82475503620492e-4
TENSION_PLASTIC = 1.88140523638028e-4
COMPRESSION_COHESION = 28.4472458689923

# Компоненты эффективной упругой пробы: xx, yy, zz, xy, xz, yz, в МПа.
# Сняты от двух принятых историй в семи отклонённых пробах A2.
PROBES = {
    6: (3.93684800207994, 3.9520981347758, 3.90372801185138,
        0.0705291743711791, 0.302299361384909, -0.302040159837977),
    7: (3.93535327899118, 3.95056396412467, 3.9023011251657,
        0.0703708052692076, 0.300958029373394, -0.300699887521766),
    8: (3.92467925049125, 3.93958940065286, 3.8918311869306,
        0.0693831070959353, 0.296958120359973, -0.296698589602584),
    9: (3.92367765723099, 3.93856022884774, 3.89101757736516,
        0.0691986571098343, 0.293084672928512, -0.292830466123371),
    10: (3.92141926458198, 3.93614058779922, 3.89303000661829,
         0.066500232762178, 0.210392652140264, -0.210193935499326),
    12: (3.91754565725314, 3.9209229697028, 3.91178456183784,
         0.0141653818281136, 0.0552028969616966, -0.0551704164620378),
    13: (3.91699893410672, 3.92036002442804, 3.91135673389847,
         0.014052838041902, 0.0528294903633997, -0.0527969147791552),
}


def invariants(stress):
    """Вычисляет след, интенсивность и крайние главные значения девиатора."""
    trace = sum(stress[:3])
    a, b, c = (value - trace / 3 for value in stress[:3])
    xy, xz, yz = stress[3:]
    q = sqrt(1.5 * (a*a + b*b + c*c) + 3 * (xy*xy + xz*xz + yz*yz))
    scale = q / 3
    determinant = a*b*c + 2*xy*xz*yz - a*yz*yz - b*xz*xz - c*xy*xy
    argument = max(-1, min(1, determinant / (2 * scale**3)))
    angle = acos(argument) / 3
    return trace, q, 2 * scale * cos(angle), 2 * scale * cos(angle + 2*pi/3)


def required_dilation(volume_rhs, radius_part):
    """Решает d√((a·d)² + b²) = H по d >= 0."""
    b_squared = radius_part**2
    radical = hypot(b_squared, 2 * ECCENTRIC_OFFSET * volume_rhs)
    return sqrt(2 * volume_rhs**2 / (b_squared + radical))


def prescribed_dilation(tension_plastic, confinement):
    """Вычисляет угол из растягивающей истории при трёх положительных напряжениях."""
    fracture_energy = 0.073 * FC**0.18
    consumed = -expm1(-FT * 125 * tension_plastic / fracture_energy)
    mixed = min(1, 21 * consumed)
    friction_peak = sin(PHI * confinement)
    dilatancy_peak = sin(PSI * confinement)
    fixed_volume = (friction_peak - dilatancy_peak) / (1 - friction_peak * dilatancy_peak)
    friction = 2 * sqrt(mixed) / (1 + mixed) * friction_peak
    if friction <= fixed_volume:
        return 0.0
    dilatancy = (friction - fixed_volume) / (1 - friction * fixed_volume)
    return dilatancy / sqrt(1 - dilatancy*dilatancy)


def non_tensile_yield_upper(q_trial, high_deviator, low_deviator):
    """Ограничивает F сверху для возврата хотя бы с одним неположительным напряжением."""
    fc_initial = 0.6 * FC
    peak_strain = 0.002
    damage = 0.423 * (1 - FC / (peak_strain * MODULUS))
    plastic_peak = peak_strain - FC / ((1 - damage) * MODULUS)
    compression = COMPRESSION_PLASTIC - low_deviator / SHEAR2
    if compression >= plastic_peak:
        raise ValueError("Граница применима на участке упрочнения при сжатии")
    degradation = -log(1 - damage) / plastic_peak
    cohesion_max = (fc_initial + (FC - fc_initial) * compression / plastic_peak) * exp(degradation * compression)
    beta_max = cohesion_max / FT * (1 - ALPHA) - (1 + ALPHA)
    numerator = q_trial - 3*ALPHA*low_deviator + beta_max * (high_deviator - low_deviator)
    return numerator / (1 - ALPHA) - COMPRESSION_COHESION


def verify_probe(stress, committed_tension):
    """Исключает корень исходной модели во всём диапазоне возврата конкретной пробы."""
    trace, q_trial, high, low = invariants(stress)
    beta = COMPRESSION_COHESION / FT * (1 - ALPHA) - (1 + ALPHA)
    slope = ALPHA + beta / 3
    apex = COMPRESSION_COHESION * (1 - ALPHA) / slope
    trace_drop = (q_trial + beta * high) / slope
    min_trace = apex - trace_drop
    yield_upper = non_tensile_yield_upper(q_trial, high, low)
    if trace <= apex or min_trace / 3 + low <= 0 or yield_upper >= 0:
        raise ValueError("Проба выходит за область доказанной растягивающей скобки")

    radius_part = sqrt(2 / 3) * q_trial
    initial_rhs = (trace - apex) * SHEAR2 / BULK9
    scale = required_dilation(initial_rhs, radius_part)
    decay = min_trace**2 / (FC * q_trial)
    # При q > 0 верхняя граница d_закона/d_возврата растёт с f до f=1.
    log_ratio = log(tan(PSI) / scale) - min_trace / FC - decay

    def volumes(factor):
        """Возвращает нормальную пластическую добавку и историю растяжения."""
        volume = (trace - apex + trace_drop * factor) / BULK9
        tension = committed_tension + volume + (1 - factor) * high / SHEAR2
        return volume, tension

    def return_dilation(factor):
        """Получает требуемую дилатансию на поверхности текучести."""
        volume, _ = volumes(factor)
        return required_dilation(volume * SHEAR2 * factor / (1 - factor), radius_part * factor)

    cutoff = 1e-12 * FC / q_trial
    if log_ratio >= 0:
        def prescribed_log_bound(factor):
            """Ограничивает дилатансию сверху по положительному следу."""
            current_trace = apex - trace_drop * factor
            return log(tan(PSI)) - current_trace / FC - current_trace**2 / (FC * q_trial * factor)

        first = 0.0001
        middle = (1 + first) / 2
        if cutoff >= first or (apex - trace_drop * first)**2 / (FC * q_trial * first) <= 0.5:
            raise ValueError("Первая граница не покрывает монотонный приосевой интервал")
        # На двух последних интервалах обе границы растут. Сопоставляем
        # предписанный максимум справа с требуемым минимумом слева.
        log_ratio = max(prescribed_log_bound(first) - log(scale) - 0.5 * log(first),
                        prescribed_log_bound(middle) - log(return_dilation(first)),
                        prescribed_log_bound(1) - log(return_dilation(middle)))
    cutoff_required = return_dilation(cutoff)  # Требуемая дилатансия монотонно растёт с f.
    minimum_tension = min(volumes(0)[1], volumes(cutoff)[1])
    cutoff_prescribed = prescribed_dilation(minimum_tension, 1)
    if decay < 0.5 or log_ratio >= 0 or cutoff_required >= cutoff_prescribed:
        raise ValueError("Безкорневой интервал исходного закона не подтверждён")

    return {
        "trial_trace_mpa": trace,
        "trial_q_mpa": q_trial,
        "apex_trace_mpa": apex,
        "non_tensile_yield_upper_mpa": yield_upper,
        "prescribed_over_required_upper": exp(log_ratio),
        "isotropic_branch_required_upper": cutoff_required,
        "isotropic_branch_prescribed_lower": cutoff_prescribed,
    }


if __name__ == "__main__":
    print(json.dumps({attempt: verify_probe(stress, EARLY_TENSION_PLASTIC if attempt < 12 else TENSION_PLASTIC)
                      for attempt, stress in PROBES.items()},
                     indent=2, allow_nan=False))
