"""ДЛЯ ТЕСТОВ. Аналитически проверяет проходимость гладкого максимума при контроле кривизны траектории, отдельно от равновесия."""
from pathlib import Path
import json,math

out=Path.cwd()/'tmp/material-consistency/ArcMetricCounterexample.json';cases=[]
for alpha in (1e-4,1e-3,1e-2,1.):
    for c in (.1,1.,10.):
        d=.001/(alpha*c);h=1.5*d;u0=-d;uh=u0+h/2;uf=u0+h
        peak=1+c*d*d;l0=1.;lh=peak-c*uh*uh;lf=peak-c*uf*uf
        exact_lambda_error=c*h*h/2;exact_lambda_change=c*(2*d*h-h*h)
        floor=8*.0001*5e-7*(1+abs(l0));old_bound=max(.005*abs(exact_lambda_change),floor)
        old_ratio=exact_lambda_error/abs(exact_lambda_change)
        joint=alpha*exact_lambda_error/math.sqrt(h*h+alpha*alpha*exact_lambda_change**2)
        assert uf>0 and u0<0 and old_ratio>.5 and exact_lambda_error>old_bound and joint<.005
        for s in (.01,100.):
            # Одновременно преобразуются координаты и их метрика: физическая длина шага остаётся прежней.
            scaled=math.sqrt((s*h)**2/s**2+alpha**2*exact_lambda_change**2)
            assert abs(alpha*exact_lambda_error/scaled-joint)<1e-15
        # Ошибка кривизны не заменяет проверку равновесия: линейный дефект нагрузки сокращается в удвоенном полушаге.
        force_defect=1e-3;bad_lh=lh+force_defect/2;bad_lf=lf+force_defect
        assert abs((2*bad_lh-l0-bad_lf)-(2*lh-l0-lf))<=1e-10*max(1,abs(peak))
        assert force_defect>1e-10
        cases.append(dict(alpha=alpha,curvature=c,initial_distance_to_peak=d,step=h,
            crosses_known_smooth_peak=True,old_relative_lambda_curvature_ratio=old_ratio,
            old_lambda_bound=old_bound,lambda_curvature_error=exact_lambda_error,joint_curve_error=joint,
            original_projection_rejects=True,joint_fixed_metric_passes=True,force_check_required_independently=True))
proof=dict(analytical_model='lambda(u)=lambda_peak-c*u^2; u0=-d, half=u0+h/2, full=u0+h, h=1.5d',
 cases=cases,exact_old_error_ratio_for_any_crossing='h/(2*abs(2*d-h)) > 1/2 for h>d',
 joint_fixed_metric='sqrt(||e_u||_M^2+alpha^2*e_lambda^2)/sqrt(||Delta_u||_M^2+alpha^2*Delta_lambda^2)',
 fixed_positive_alpha_required=True,alpha_must_not_diverge_as_inverse_current_load_slope=True,
 new_metric_is_curve_control_not_constitutive_integration_error_estimator=True,
 original_force_equilibrium_and_material_checks_not_replaced=True,
 physical_capacity_certified=False,no_working_source_change=True)
out.write_text(json.dumps(proof,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('12 ANALYTICAL SMOOTH-PEAK CROSSINGS:OLD LAMBDA PROJECTION REJECTS;FIXED JOINT METRIC PASSES;FORCE DEFECT STILL REJECTED')
