"""Проверяет объёмную модель HardShear в OpenSeesPy (Н, мм, МПа).

Закреплена нижняя грань. Поперечные силы создают градиент моментов вдоль
призмы. В контрольной плоскости поддерживается заданная траектория шести
усилий. Остановка продолжения сохраняется отдельно от достижения предела.
"""
import argparse
import json
import os
import math
from itertools import product
from pathlib import Path
from time import perf_counter

import openseespy.opensees as ops

TARGET = (-28000., 510000., 180000., 1500000., 17000., -7100.)
X = [-90., -76., -38., 0., 38., 76., 90.]
Y = [-60., -46., -30., 0., 30., 46., 60.]
AREA, IX, IY = 180.*120., 180.*120.**3/12., 120.*180.**3/12.
E, NU, ES, POST = 30000., .2, 200000., 1500.
GAUSS = [-1/math.sqrt(3), 1/math.sqrt(3)]
CORNERS = [(0,0), (1,0), (1,1), (0,1)]
CROSS_TOL = 0.03  # N, or N.mm divided by 180 mm


def scale(factor, values):
    return [factor * value for value in values]


def add(target, values):
    for i, value in enumerate(values):
        target[i] += value


def wrench(point, force, ref):
    x, y, z = point
    fx, fy, fz = force
    return [fz, y*fz-(z-ref)*fy, x*fz-(z-ref)*fx, x*fy-y*fx, fx, fy]


def error(actual, expected):
    return max(abs(a-b)/(180. if i in (1,2,3) else 1.)
               for i,(a,b) in enumerate(zip(actual,expected,strict=True)))


def envelope(yield_stress, peak_stress, peak_plastic, damage_rate, energy):
    softening_rate = peak_stress/(energy - (peak_stress+yield_stress)*peak_plastic/2)
    knots = [peak_plastic*i/40 for i in range(40)] if peak_plastic else []
    # Tail knots resolve a 5% step in the softening exponent down to 1e-6 nominal strength.
    knots += [peak_plastic+i/20/softening_rate for i in range(277)]
    rows = [(0., 0., 0.)]
    for plastic in knots:
        stress = (yield_stress+(peak_stress-yield_stress)*plastic/peak_plastic
                  if plastic < peak_plastic else peak_stress*math.exp(-softening_rate*(plastic-peak_plastic)))
        d = -math.expm1(-damage_rate*plastic)
        effective = stress*math.exp(damage_rate*plastic)
        eps = plastic+effective/E
        if eps>.1 or d>1.-1.e-12:
            break
        rows.append((eps, stress, d))
    for prev,curr in zip(rows,rows[1:]):
        assert curr[0]>prev[0] and curr[1]>0 and prev[2]<=curr[2]<1
    return rows


def concrete(variant):
    if variant == "elastic":
        ops.nDMaterial("ElasticIsotropic",1,E,NU)
    elif variant == "default":
        ops.nDMaterial("ASDConcrete3D",1,E,NU,"-fc",30.,"-ft",3.,"-tangent")
    else:
        dc = .423*(1.-30./(.002*E))
        plastic = .002-30./((1.-dc)*E)
        tension = envelope(3.,3.,0.,3.*125./.12,.12/125.)
        compression = envelope(16.,30.,plastic,-math.log1p(-dc)/plastic,80./125.)
        args = ["ASDConcrete3D",1,E,NU]
        for prefix,rows in (("T",tension),("C",compression)):
            for j,field in enumerate(("e","s","d")):
                args.extend(["-"+prefix+field,*(r[j] for r in rows)])
        ops.nDMaterial(*args,"-tangent")
        return tension, compression


def steel(variant):
    if variant == "elastic":
        ops.uniaxialMaterial("Elastic",2,ES)
        ops.uniaxialMaterial("Elastic",3,ES)
    else:
        def backbone(fy, sign):
            yield_strain = fy/ES
            return [sign*item for strain in (yield_strain,.025,.10)
                    for item in (fy+POST*(strain-yield_strain),strain)]
        ops.uniaxialMaterial("Hysteretic",2,*backbone(400.,1),*backbone(350.,-1),1.,1.,0.,0.,0.)
        ops.uniaxialMaterial("Steel01",3,240.,ES,POST/ES)


def mesh_axis(values, refine):
    return [a+(b-a)*i/refine for a,b in zip(values,values[1:]) for i in range(refine)]+[values[-1]]


class Prism:
    def __init__(self,length,h,variant,refine,control):
        if length <= 0 or h <= 0 or not math.isclose(length / h, round(length / h)):
            raise ValueError("Length must be a positive multiple of dz")
        if not 0 < control < length or refine < 1:
            raise ValueError("Control cut must be internal and refinement must be positive")
        self.x,self.y = mesh_axis(X,refine),mesh_axis(Y,refine)
        self.nx,self.ny,self.nz = len(self.x)-1,len(self.y)-1,round(length/h)
        self.z = [i*h for i in range(self.nz+1)]
        self.mid = self.z.index(control)
        self.coords = {}
        self.bricks = [[] for _ in range(self.nz)]
        self.long = [[] for _ in range(self.nz)]
        self.bars = {}
        self.planes = [[] for _ in self.z]
        ops.wipe()
        ops.model("basic","-ndm",3,"-ndf",3)
        reference = concrete(variant)
        steel(variant)
        for k,j,i in product(range(self.nz+1),range(self.ny+1),range(self.nx+1)):
            p = self.x[i],self.y[j],self.z[k]
            tag = self.node(i,j,k)
            self.coords[tag] = p
            self.planes[k].append(tag)
            ops.node(tag,*p)
            if k==0:
                ops.fix(tag,1,1,1)
        self.next = 1
        for k,j,i in product(range(self.nz),range(self.ny),range(self.nx)):
            tags = [self.node(i+dx,j+dy,k+dz) for dz in (0,1) for dx,dy in CORNERS]
            ops.element("stdBrick",self.next,*tags,1)
            self.bricks[k].append((self.next,tags))
            self.next += 1
        self.installed_adjustment = None
        if reference:
            self.installed_adjustment = []
            for side,rows in zip(("T","C"),reference,strict=True):
                tag = self.bricks[self.mid-1][0][0]
                e,s,q = [ops.eleResponse(tag,"material",1,side+suffix) for suffix in ("e","s","q")]
                errors = [max(abs(a-b) for a,b in zip(actual,expected,strict=True))
                          for actual,expected in ((e,[r[0] for r in rows]),(s,[r[1] for r in rows]),
                                                   (q,[r[1]/(1-r[2]) for r in rows]))]
                if max(errors)>1.e-7:
                    raise RuntimeError(("installed concrete envelope altered",side,errors))
                self.installed_adjustment.append(errors)
        ix = self.x.index(0.)
        for k in range(self.nz):
            for y in (-30.,30.):
                iy = self.y.index(y)
                self.long[k].append(self.bar(self.node(ix,iy,k),self.node(ix,iy,k+1),2,12.))
        left,right = self.x.index(-76.),self.x.index(76.)
        lo,hi = self.y.index(-46.),self.y.index(46.)
        loop = ([(i,lo) for i in range(left,right)]+[(right,j) for j in range(lo,hi)]
                +[(i,hi) for i in range(right,left,-1)]+[(left,j) for j in range(hi,lo,-1)])
        self.hoop_levels = list(range(50,round(length),100))
        assert control not in self.hoop_levels
        for level in self.hoop_levels:
            k = self.z.index(level)
            nodes = [self.node(i,j,k) for i,j in loop]
            for a,b in zip(nodes,nodes[1:]+nodes[:1],strict=True):
                self.bar(a,b,3,8.)
        self.loads = self.apply_loads()

    def node(self,i,j,k):
        return 1+(k*(self.ny+1)+j)*(self.nx+1)+i

    def bar(self,a,b,material,diameter):
        tag = self.next
        self.next += 1
        ops.element("truss",tag,a,b,math.pi*diameter**2/4,material)
        self.bars[tag] = (a,b,material)
        return tag

    def resultant(self,values,ref):
        total = [0.]*6
        for tag,force in values.items():
            add(total,wrench(self.coords[tag],force,ref))
        return total

    def apply_loads(self):
        ops.timeSeries("Linear",1)
        ops.pattern("Plain",1,1)
        result = {tag:[0.]*3 for tag in self.planes[-1]}
        dist = self.z[-1]-self.z[self.mid]
        n,mx,my,t,qx,qy = TARGET
        mx,my = mx+dist*qy,my+dist*qx
        for _,nodes in self.bricks[-1]:
            top = nodes[4:]
            x0,y0,_ = self.coords[top[0]]
            x1,y1,_ = self.coords[top[2]]
            jac = (x1-x0)*(y1-y0)/4
            for r,s in product(GAUSS,repeat=2):
                x,y = (x0+x1+r*(x1-x0))/2,(y0+y1+s*(y1-y0))/2
                forces = [qx/AREA-t*y/(IX+IY),qy/AREA+t*x/(IX+IY),n/AREA+mx*y/IX+my*x/IY]
                for tag,(i,j) in zip(top,CORNERS,strict=True):
                    shape = (1+(2*i-1)*r)*(1+(2*j-1)*s)/4
                    add(result[tag],scale(shape*jac,forces))
        for tag,force in result.items():
            ops.load(tag,*force)
        applied = self.resultant(result,self.z[self.mid])
        if error(applied,TARGET)>1.e-8:
            raise RuntimeError(("wrench at control plane",applied,TARGET))
        return result

    def cut(self,layer,lower):
        end = int(lower)
        sign = 2*end-1
        ref = self.z[layer+end]
        total = [0.]*6
        for tag,nodes in self.bricks[layer]:
            force = ops.eleForce(tag)
            assert len(force)==24
            for i in range(4*end,4*end+4):
                add(total,wrench(self.coords[nodes[i]],scale(sign,force[3*i:3*i+3]),ref))
        for tag in self.long[layer]:
            force = ops.eleForce(tag)
            assert len(force)==6
            node = self.bars[tag][end]
            add(total,wrench(self.coords[node],scale(sign,force[3*end:3*end+3]),ref))
        return total

    def field(self,variant):
        target = scale(ops.getTime(),TARGET)
        ops.reactions()
        base = self.resultant({node:scale(-1,ops.nodeReaction(node)) for node in self.planes[0]},self.z[self.mid])
        low,up = self.cut(self.mid-1,True),self.cut(self.mid,False)
        errors = [error(expected,target) for expected in (base,low,up)]
        if max(errors)>CROSS_TOL:
            raise RuntimeError(("accepted step violates equilibrium",ops.getTime(),errors,target,base,low,up))
        strains = {}
        for tag,(a,b,kind) in self.bars.items():
            start,end = self.coords[a],self.coords[b]
            vector = [y-x for x,y in zip(start,end,strict=True)]
            delta = [y-x for x,y in zip(ops.nodeDisp(a),ops.nodeDisp(b),strict=True)]
            eps = sum(v*u for v,u in zip(vector,delta,strict=True))/sum(v*v for v in vector)
            actual = ops.eleResponse(tag,"deformations")
            if len(actual)!=1 or not math.isclose(actual[0],eps*math.dist(start,end),abs_tol=1.e-9):
                raise RuntimeError(("truss deformation",tag,eps,actual))
            z = (start[2]+end[2])/2
            label = "long" if kind==2 else "hoop"
            limit = ((.025 if eps>=0 else .018) if kind==2 else .02)
            demand = dict(util=abs(eps)/limit,strain=eps,tag=tag,start=start,end=end,z=z)
            for key in [label] + ([label+"_near_cut"] if abs(z-self.z[self.mid])<=50 else []):
                if key not in strains or demand["util"]>strains[key]["util"]:
                    strains[key] = demand
        damage = None
        damage_max = None
        if variant!="elastic":
            damage = [0.,0.]
            damage_max = [{"value": 0.}, {"value": 0.}]
            for layer in range(self.nz):
                for tag,_ in self.bricks[layer]:
                    for point in range(1,9):
                        values = ops.eleResponse(tag,"material",point,"damage")
                        assert len(values)==2 and all(map(math.isfinite,values))
                        if layer in (self.mid-1,self.mid):
                            damage = [max(a,b) for a,b in zip(values,damage,strict=True)]
                        for i,value in enumerate(values):
                            if value > damage_max[i]["value"]:
                                damage_max[i] = dict(value=value, element=tag, gauss_point=point,
                                                     z_interval_mm=self.z[layer:layer+2])
        return dict(factor=ops.getTime(),equilibrium_error=errors,lower_cut=low,upper_cut=up,
                    base_at_cut=base,steel=strains,adjacent_damage=damage,prism_damage_max=damage_max,
                    top_axis_displacement=ops.nodeDisp(self.node(self.x.index(0),self.y.index(0),self.nz)))

    def displacements(self):
        return [v for node in self.coords for v in ops.nodeDisp(node)]


def analyze(prism,opts):
    ops.constraints("Transformation")
    ops.numberer("RCM")
    ops.system("UmfPack")
    ops.test("NormUnbalance",2.e-4,35)
    ops.algorithm("Linear" if opts.variant=="elastic" else "Newton")
    ops.integrator("LoadControl",1.)
    ops.analysis("Static")
    output = dict(opensees=ops.version(),variant=opts.variant,section_target=TARGET,
                  length=opts.length,z_control=opts.cut,mesh=[prism.nx,prism.ny,prism.nz],
                  hoop_planes=prism.hoop_levels,installed_adjustment=prism.installed_adjustment,
                  cut_check_error_N=CROSS_TOL,applied_at_free_face=prism.resultant(prism.loads,prism.z[-1]),
                  states=[],rejected=[],termination=None)
    start = perf_counter()

    def save():
        output["elapsed_s"] = perf_counter()-start
        temporary = opts.out.with_suffix(".partial")
        temporary.write_text(json.dumps(output,indent=2,allow_nan=False),encoding="utf-8")
        os.replace(temporary,opts.out)

    def accept(mode):
        row = prism.field(opts.variant)
        row["mode"] = mode
        row["iterations"] = ops.testIter()
        output["states"].append(row)
        save()
        if len(output["states"])%10==0 or abs(row["factor"]-1)<1.e-9:
            print("accept",opts.variant,opts.length,row["factor"],row["adjacent_damage"],flush=True)
        return row

    if opts.variant=="elastic":
        if ops.analyze(1):
            raise RuntimeError("elastic sign calibration did not converge")
        accept("Linear")
        output["termination"]="elastic calibration"
        save()
        return

    dt = opts.step
    latest_increment = None
    ceiling = min(opts.maximum,opts.continuation_start)
    marks = [mark for mark in (1.,1.10263192386604) if mark < ceiling] + [ceiling]
    while ops.getTime() < ceiling-1.e-9:
        base = ops.getTime()
        before = prism.displacements()
        cap = min(mark for mark in marks if mark>base+1.e-9)
        trial = min(dt,cap-base)
        ops.integrator("LoadControl",trial)
        accepted = False
        for algorithm,args in (("Newton",()),("NewtonLineSearch",("-type","Bisection")),("KrylovNewton",())):
            ops.test("NormUnbalance",2.e-4,opts.iterations)
            ops.algorithm(algorithm,*args)
            code = ops.analyze(1)
            if not code:
                accepted = True
                break
            if not math.isclose(ops.getTime(),base,abs_tol=1.e-13):
                raise RuntimeError(("failed step did not restore factor",base,ops.getTime()))
            output["rejected"].append(dict(mode=algorithm,factor=base+trial,code=code))
        if accepted:
            state = accept(algorithm)
            latest_increment = [new-old for new,old in zip(prism.displacements(),before,strict=True)]
            if max(v["util"] for v in state["steel"].values())>=1.:
                output["termination"]="steel strain criterion crossed"
                break
            dt = min(opts.step,1.25*dt)
        else:
            dt = trial/2
            if dt<opts.minimum:
                output["termination"]="load control continuation stopped"
                break
    if output["termination"] is None:
        output["termination"] = ("factor ceiling reached" if ceiling == opts.maximum
                                 else "start of prescribed arc continuation")
    save()

    if not opts.arc or latest_increment is None or output["termination"] not in (
            "load control continuation stopped","start of prescribed arc continuation"):
        return
    ops.algorithm("Newton")
    # Keep the active integrator between converged steps. ArcLength preserves the
    # sign of its last increment internally to follow a descending load branch.
    length = (opts.arc_length or max(.0001,min(.04,math.hypot(*latest_increment))))
    previous_increment = latest_increment
    ops.integrator("ArcLength",length,0.)
    consecutive_drops = 0
    for _ in range(opts.arc):
        base = ops.getTime()
        before = prism.displacements()
        ops.test("NormUnbalance",2.e-4,opts.iterations)
        code = ops.analyze(1)
        if code:
            if not math.isclose(ops.getTime(),base,abs_tol=1.e-13):
                raise RuntimeError(("failed arc step did not restore factor",base,ops.getTime()))
            output["rejected"].append(dict(mode="ArcLength",length=length,factor=base,code=code))
            if consecutive_drops:
                output["termination"] = "arc stopped on descending trial branch"
                break
            length /= 2
            if length < 1.e-5:
                output["termination"] = "arc continuation stopped"
                break
            ops.integrator("ArcLength",length,0.)
            continue
        increment = [new-old for new,old in zip(prism.displacements(),before,strict=True)]
        cosine = (math.fsum(a*b for a,b in zip(increment,previous_increment,strict=True))
                  / (math.hypot(*increment)*math.hypot(*previous_increment)))
        state = accept("ArcLength")
        state["arc_length"] = length
        state["arc_cosine_previous"] = cosine
        previous_increment = increment
        if cosine <= 0.:
            output["termination"] = "arc reversed displacement direction"
            break
        consecutive_drops = consecutive_drops+1 if state["factor"]<base-1.e-7 else 0
        if max(v["util"] for v in state["steel"].values())>=1.:
            output["termination"]="steel strain criterion crossed on arc"
            break
        if consecutive_drops>=5:
            output["termination"]="traced five consistently oriented descending arc steps"
            break
    else:
        output["termination"] = "arc step budget exhausted"
    save()


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--length",type=float,default=400.)
    p.add_argument("--cut",type=float,default=200.)
    p.add_argument("--dz",type=float,default=25.)
    p.add_argument("--refine",type=int,default=1)
    p.add_argument("--maximum",type=float,default=3.)
    p.add_argument("--step",type=float,default=.05)
    p.add_argument("--minimum",type=float,default=.0003)
    p.add_argument("--iterations",type=int,default=30)
    p.add_argument("--arc",type=int,default=40)
    p.add_argument("--arc-length",type=float)
    p.add_argument("--continuation-start",type=float,default=float("inf"))
    p.add_argument("--variant",choices=["elastic","default","matched"],default="matched")
    p.add_argument("--out",type=Path,required=True)
    opts=p.parse_args()
    analyze(Prism(opts.length,opts.dz,opts.variant,opts.refine,opts.cut),opts)
    print("saved",opts.out,flush=True)


if __name__=="__main__":
    main()
