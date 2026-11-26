"""ДЛЯ ТЕСТОВ. Строит научные графики фактического участка бетона и упругой стали без подмены пространственного закона одноосной кривой."""
from pathlib import Path
from hashlib import sha256
import json,numpy as np
import matplotlib
matplotlib.use('Agg')
from matplotlib import pyplot as plt
repo=Path.cwd();out=repo/'docs/reports/figures/MaterialState_2026-11-26'
data=json.loads((out/'MaterialDiagramData.json').read_bytes())
plt.rcParams.update({'font.family':'DejaVu Sans','font.size':18,'axes.titlesize':20,'axes.labelsize':18,
    'xtick.labelsize':16,'ytick.labelsize':16,'legend.fontsize':15,'svg.fonttype':'none',
    'axes.spines.top':False,'axes.spines.right':False,'axes.grid':True,'grid.alpha':.22})
blue='#006BA4';orange='#D97706';red='#C62828';gray='#626262'
c=data['concrete'];history=c['trajectory'];local=c['local'];cur=c['current'];event=c['event']
x=np.array([p['epsilon1']*1e3 for p in history]);y=np.array([p['sigma1_MPa'] for p in history])
f=np.array([p['fraction'] for p in local]);sx=np.array([(p['epsilon1']-cur['epsilon1'])*1e9 for p in local]);sy=np.array([(p['sigma1_MPa']-cur['sigma1_MPa'])*1e6 for p in local])
assert len(x)==25 and np.all(np.diff(f)>0) and abs(sx[0])<1e-12 and abs(sy[0])<1e-12
ex=(event['epsilon1']-cur['epsilon1'])*1e9;ey=(event['sigma1_MPa']-cur['sigma1_MPa'])*1e6
fig,axes=plt.subplots(3,1,figsize=(8.6,16),layout='constrained')
fig.suptitle('Бетон: элемент №11, станция z/L = −0,001',fontsize=21)
ax=axes[0];ax.plot(x,y,color=blue,lw=1.5,marker='.',ms=5,label='25 состояний в порядке расчёта')
ax.scatter([cur['epsilon1']*1e3],[cur['sigma1_MPa']],marker='x',color=red,s=140,lw=3,zorder=5,label='Итерация остановки')
ax.set(xlabel='Наибольшая главная деформация ε₁, ‰',ylabel='Наибольшее главное σ₁, МПа',
    title='Траектория точки; не одноосная диаграмма')
ax.set_xlim(0,max(x.max(),cur['epsilon1']*1e3)*1.12);ax.set_ylim(0,max(y.max(),cur['sigma1_MPa'])*1.12)
ax.annotate('Место остановки',xy=(cur['epsilon1']*1e3,cur['sigma1_MPa']),xytext=(.32,.67),textcoords='axes fraction',
    arrowprops={'arrowstyle':'->','color':red,'lw':1.5},color=red,fontsize=17)
ax.legend(loc='lower right',frameon=False)
ax=axes[1]
ax.plot(x[14:18],y[14:18],color=orange,lw=2,marker='o',ms=5,label='15 → 18: снижение σ₁')
ax.plot(x[17:25],y[17:25],color=blue,lw=2,marker='o',ms=5,label='18 → 25: повторный рост σ₁')
for a,b,col in ((16,17,orange),(18,19,blue),(21,22,blue)):
    ax.annotate('',xy=(x[b],y[b]),xytext=(x[a],y[a]),arrowprops={'arrowstyle':'->','color':col,'lw':2})
for i,offset in ((14,(8,8)),(16,(8,-20)),(17,(-5,-24)),(24,(8,8))):
    ax.annotate(str(i+1),xy=(x[i],y[i]),xytext=offset,textcoords='offset points',fontsize=14)
ax.set(xlabel='Наибольшая главная деформация ε₁, ‰',ylabel='Наибольшее главное σ₁, МПа',
    title='Увеличение «крючка»: разгрузка и возврат')
ax.text(.37,.10,'λ₁₇ = 1,1040 → λ₁₈ = 1,0637\n→ λ₂₅ = 1,1035',transform=ax.transAxes,fontsize=14)
ax.margins(x=.13,y=.18);ax.legend(loc='upper left',frameon=False,fontsize=14)
ax=axes[2];ax.plot(sx,sy,color=gray,lw=2.3,label='Поля вдоль пробной поправки')
ax.scatter([0],[0],color=red,marker='x',s=140,lw=3,zorder=5,label='Итерация остановки')
ax.scatter([ex],[ey],color=orange,s=90,zorder=5,label='Возврат на пластическую ветвь')
ax.axvline(ex,color=orange,lw=1.1,ls='--',alpha=.7)
ax.set(xlabel='Приращение Δε₁, 10⁻⁹',ylabel='Приращение Δσ₁, 10⁻⁶ МПа',title='Увеличение перехода в точке остановки')
ax.legend(loc='best',frameon=False)
fig.savefig(out/'ConcreteActualState.png',dpi=180)
fig.savefig(out/'ConcreteActualState.svg')
plt.close(fig)

fig,axes=plt.subplots(2,1,figsize=(8.6,10.8),layout='constrained')
fig.suptitle('Арматура: текущие точки на диаграмме',fontsize=21)
ax=axes[0];ax.plot([0,2000],[0,400],color=gray,lw=2.2)
ax.scatter([2000],[400],color=blue,s=85,marker='o',label='Продольная: начало текучести 400 МПа')
ax.scatter([1200],[240],color=orange,s=100,marker='^',label='Поперечная: начало текучести 240 МПа')
for s,col,mark in zip(data['steel'],(blue,orange),('o','^')):
    ax.scatter([s['epsilon']*1e6],[s['sigma_MPa']],color=col,s=70,marker=mark,zorder=5)
ax.axvspan(0,120,color=blue,alpha=.09)
ax.set(xlim=(0,2150),ylim=(0,445),xlabel='Относительная деформация εₛ, 10⁻⁶',ylabel='Напряжение σₛ, МПа',title='Упругая ветвь до начала текучести')
ax.legend(loc='lower right',frameon=False,fontsize=14)
ax.annotate('Текущие точки',xy=(90,18),xytext=(.34,.35),textcoords='axes fraction',
    arrowprops={'arrowstyle':'->','color':red,'lw':1.5},color=red,fontsize=17)
ax=axes[1];ax.plot([0,120],[0,24],color=gray,lw=2.2,label='Eₛ = 200 000 МПа')
for s,col,mark in zip(data['steel'],(blue,orange),('o','^')):
    xx=s['epsilon']*1e6;yy=s['sigma_MPa'];ax.scatter([xx],[yy],color=col,s=110,marker=mark,zorder=5,
        label=s['label']+': '+format(yy,'.2f').replace('.',',')+' МПа')
ax.set(xlim=(0,120),ylim=(0,26),xlabel='Относительная деформация εₛ, 10⁻⁶',ylabel='Напряжение σₛ, МПа',title='Увеличение текущего упругого участка')
ax.legend(loc='upper left',frameon=False)
fig.savefig(out/'SteelActualState.png',dpi=180)
fig.savefig(out/'SteelActualState.svg')
plt.close(fig)
(out/'ConcreteActualState.svg').write_text('\n'.join(line.rstrip() for line in (out/'ConcreteActualState.svg').read_text('utf-8').splitlines())+'\n',encoding='utf-8')
(out/'SteelActualState.svg').write_text('\n'.join(line.rstrip() for line in (out/'SteelActualState.svg').read_text('utf-8').splitlines())+'\n',encoding='utf-8')
(out/'Plot-TestMaterialState.py').write_bytes(Path(__file__).read_bytes())
proof=dict(data_sha256=sha256((out/'MaterialDiagramData.json').read_bytes()).hexdigest().upper(),
    matplotlib_version=matplotlib.__version__,concrete_is_actual_multiaxial_trajectory_not_uniaxial_envelope=True,
    steel_only_elastic_segments_drawn=True,plotted_point_is_actual934_record3=True,
    failed_iteration_not_equilibrium_or_failure_capacity=True,
    current_concrete_epsilon1=cur['epsilon1'],current_concrete_sigma1_MPa=cur['sigma1_MPa'],
    event_delta_epsilon1=ex*1e-9,event_delta_sigma1_MPa=ey*1e-6,
    physical_capacity_certified=False)
(out/'PlotProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PLOTS WRITTEN;CONCRETE PROJECTION AND ACTUAL ELASTIC STEEL;MPL',matplotlib.__version__)
