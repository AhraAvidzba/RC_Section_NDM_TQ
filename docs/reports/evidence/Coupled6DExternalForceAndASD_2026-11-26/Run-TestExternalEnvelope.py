"""ДЛЯ ТЕСТОВ. SciPy решает фактическую систему VBA через отдельный собственный Excel; рабочая книга не изменяется."""

from argparse import ArgumentParser
import ctypes
from ctypes import wintypes
import gc
from hashlib import sha256
import json
import os
from pathlib import Path
from runpy import run_path
import shutil
import site
import struct
import threading
import time
import traceback
from uuid import uuid4
from types import SimpleNamespace

site.addsitedir('D:/CodexScratch/RC_Section_NDM_TQ/CoupledEvidencePython')
import numpy as np
import scipy
from scipy.optimize import least_squares
import pythoncom
import win32com.client
import win32com.client.dynamic
import win32process
import pywintypes


def TestWrite(path,value):
    """Сохраняет UTF-8 и точные окончания строк независимо от платформенного текстового преобразования."""
    path.write_bytes((json.dumps(value,ensure_ascii=False,indent=2,sort_keys=True,allow_nan=False)+'\n').encode('utf-8'))


def TestProcess(pid,expected=None,terminate=False):
    """Читает точное FILETIME; завершение допустимо только при совпадении времени собственного процесса."""
    api=ctypes.WinDLL('kernel32',use_last_error=True)
    api.OpenProcess.argtypes=(wintypes.DWORD,wintypes.BOOL,wintypes.DWORD);api.OpenProcess.restype=wintypes.HANDLE
    api.GetProcessTimes.argtypes=(wintypes.HANDLE,ctypes.POINTER(wintypes.FILETIME),ctypes.POINTER(wintypes.FILETIME),ctypes.POINTER(wintypes.FILETIME),ctypes.POINTER(wintypes.FILETIME))
    api.GetProcessTimes.restype=wintypes.BOOL
    api.CloseHandle.argtypes=(wintypes.HANDLE,);api.CloseHandle.restype=wintypes.BOOL
    api.TerminateProcess.argtypes=(wintypes.HANDLE,wintypes.UINT);api.TerminateProcess.restype=wintypes.BOOL
    handle=api.OpenProcess(0x1000 | (1 if terminate else 0),False,pid)
    if not handle:
        error=ctypes.get_last_error()
        return {'state':'Absent' if error==87 else 'Unreadable','error':error}
    try:
        created,exit_time,kernel,user=(wintypes.FILETIME() for _ in range(4))
        if not api.GetProcessTimes(handle,ctypes.byref(created),ctypes.byref(exit_time),ctypes.byref(kernel),ctypes.byref(user)):
            raise ctypes.WinError(ctypes.get_last_error())
        value=(created.dwHighDateTime<<32)|created.dwLowDateTime
        if expected is not None and value!=expected:return {'state':'Replaced','filetime':value}
        exited=(exit_time.dwHighDateTime<<32)|exit_time.dwLowDateTime
        if exited:return {'state':'Exited','filetime':value,'exit_filetime':exited,'start_ticks_utc':value+504911232000000000}
        if terminate:
            if expected is None:raise ValueError('Explicit owned creation time required')
            if not api.TerminateProcess(handle,124):raise ctypes.WinError(ctypes.get_last_error())
        return {'state':'TerminationRequested' if terminate else 'Alive','filetime':value,'start_ticks_utc':value+504911232000000000}
    finally:api.CloseHandle(handle)


def TestRun(root,method,seed_kind,fraction,step,max_nfev,timeout):
    """Каждый отклик и его аргумент сохраняются как Double; успех оптимизатора отделён от исходного равновесия и приёмки ветви."""
    prep=json.loads((root/'Preparation.json').read_bytes())
    for name,digest in prep['inputs'].items():assert sha256((root/name).read_bytes()).hexdigest().upper()==digest,name
    compilation=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
    assert compilation['Compiled'] and not compilation['ProductionChanged']
    assert json.loads((root/'CompileClosure.json').read_bytes())['State']=='Absent'
    source=root/'docs/regression/Diagnostic.xlsm'; source_hash=sha256(source.read_bytes()).hexdigest().upper()
    assert source_hash==compilation['WorkbookSHA256']
    output=root/'Native/ExternalResult.json';assert not output.exists()
    fixture=root/'Temp'/('External_'+uuid4().hex)/'Diagnostic.xlsm';fixture.parent.mkdir();shutil.copyfile(source,fixture)
    before=set(win32process.EnumProcesses());owner={}; excel=None;book=None;trace=None;timer=None;done=threading.Event()
    trace_lock=threading.Lock();started=time.perf_counter();result={'method':method,'seed_kind':seed_kind,
        'relative_energy_increment':fraction,'difference_relative_step':step,'maximum_base_evaluations':max_nfev,
        'scipy_version':scipy.__version__,'numpy_version':np.__version__,'source_sha256':source_hash,
        'fixture':str(fixture),'physical_capacity_certified':False,'working_external_dependency_added':False,
        'candidate_diagnostic_replay_not_branch_admission':True}
    def TestTimeout():
        """Watchdog сохраняет доказательства и касается лишь нового Excel с точным временем запуска."""
        if done.is_set():return
        with trace_lock:
            if trace is not None and not trace.closed:trace.flush()
        state=TestProcess(owner['pid'],owner['filetime'],True) if owner.get('exclusive_private_fixture') and not owner.get('protected_workbooks_detected') else {'state':'NoExclusiveOwnedExcel'}
        TestWrite(root/'Native/ExternalTimeout.json',dict(owner=owner,after=state,timeout_seconds=timeout,physical_failure=False))
        os._exit(124)
    pythoncom.CoInitialize()
    try:
        excel=win32com.client.dynamic.Dispatch(pythoncom.CoCreateInstance('Excel.Application',None,pythoncom.CLSCTX_LOCAL_SERVER,pythoncom.IID_IDispatch))
        _,pid=win32process.GetWindowThreadProcessId(int(excel.Hwnd))
        if pid in before:raise RuntimeError('Excel COM did not create a new owned process; existing application protected')
        state=TestProcess(pid);assert state['state']=='Alive'
        owner.update(pid=pid,filetime=state['filetime'],start_ticks_utc=str(state['start_ticks_utc']))
        TestWrite(root/'Native/ExternalOwnership.json',owner)
        timer=threading.Timer(timeout,TestTimeout);timer.daemon=True;timer.start()
        excel.Visible=False;excel.DisplayAlerts=False;excel.EnableEvents=False;excel.ScreenUpdating=False;excel.AutomationSecurity=1
        if excel.Workbooks.Count:raise RuntimeError('Unexpected workbooks in newly created Excel')
        book=excel.Workbooks.Open(str(fixture),0,False)
        if bool(book.ReadOnly):raise RuntimeError('Private fixture opened read-only')
        if Path(str(book.FullName)).resolve()!=fixture.resolve():raise RuntimeError('Opened fixture differs')
        owner['exclusive_private_fixture']=True;TestWrite(root/'Native/ExternalOwnership.json',owner)
        prefix="'"+str(book.Name).replace("'","''")+"'!modTestCoupledElasticState."
        boot=excel.Run(prefix+'TestStartExternalEnvelope',str(root));result['bootstrap']=str(boot)
        if 'RUNTIME ERROR:' in str(boot):raise RuntimeError(str(boot))
        result['context_open_attempts']=[]
        for context_attempt in range(3):
            try:
                opened=excel.Run(prefix+'TestOpenExternalEnvelope',float(fraction),int(seed_kind))
                result['context_open_attempts'].append(dict(attempt=context_attempt+1,outcome='returned'))
                break
            except pywintypes.com_error as error:
                result['context_open_attempts'].append(dict(attempt=context_attempt+1,outcome='COM error',error=str(error)))
                TestWrite(root/'Native/ExternalContextOpen.json',result['context_open_attempts'])
                if error.hresult!=-2147417856 or context_attempt==2:raise
                assert TestProcess(owner['pid'],owner['filetime'])['state']=='Alive'
                assert Path(str(book.FullName)).resolve()==fixture.resolve()
                time.sleep(.2) # Повторяем только восстановление того же контекста из неизменного снимка, без принятия новых полей.
        meta=np.asarray(opened,float).ravel()
        assert meta.shape==(11,);n=int(meta[0]);result['context_metadata']=meta.tolist()
        scales=np.asarray(excel.Run(prefix+'TestExternalEnvelopeScales'),float).ravel();assert scales.shape==(n,)
        result['unknown_scales']=scales.tolist()
        trace=(root/'Native/ExternalFTrace.bin').open('xb');trace.write(b'NDMEXTFTRACE0001'+struct.pack('<i',n));trace.flush()
        calls=0;jacobians=0;point_calls=0
        material_path=root/'Input/MaterialControl.json'
        material=json.loads(material_path.read_bytes()) if material_path.exists() else None
        if material is not None:result['material_control']=material
        def TestResidual(x):
            """Делегирует вычисление исходному VBA, без штрафных или искусственных значений невязки."""
            nonlocal calls
            x=np.asarray(x,float);assert x.shape==(n,) and np.all(np.isfinite(x))
            argument=win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8,tuple(float(z) for z in x))
            value=np.asarray(excel.Run(prefix+'TestEvaluateExternalEnvelope',argument),float).ravel()
            assert value.shape==(n,) and np.all(np.isfinite(value))
            if material is not None:
                if material.get('mode')=='projected':
                    mapped=np.asarray(excel.Run(prefix+'TestMapExternalProjectedControl',argument),float).ravel()
                    assert mapped.shape==(n,)
                    with (root/'Native/ExternalMappedCoordinates.bin').open('ab') as mapped_stream:
                        mapped_stream.write(struct.pack('<Q',calls+1)+mapped.astype('<f8').tobytes())
                else:
                    points=TestPointTrace(x)
                    selected=points[(points[:,0]==material['point'])&(points[:,1]==material['side'])]
                    assert len(selected)==1
                    with (root/'Native/ExternalControlRows.bin').open('ab') as control_stream:
                        control_stream.write(struct.pack('<Qdd',calls+1,float(value[-1]),float(selected[0,17])))
                    value[-1]=selected[0,17]-material['target_MPa']
            calls+=1
            with trace_lock:
                trace.write(struct.pack('<Qd',calls,time.perf_counter()-started));trace.write(x.astype('<f8').tobytes());trace.write(value.astype('<f8').tobytes());trace.flush()
            if calls%100==0:print('EXACT_VBA_F_CALLS',calls,'MAX_MPa',float(np.max(np.abs(value))),flush=True)
            return value
        def TestJacobian(x):
            """Повторяет центральный шаг нормированных координат VBA; SciPy меняет только алгоритм решения."""
            nonlocal jacobians
            jacobians+=1;matrix=np.empty((n,n))
            for i in range(n):
                delta=step*(1+abs(x[i]));a=x.copy();b=x.copy();a[i]+=delta;b[i]-=delta
                matrix[:,i]=(TestResidual(a)-TestResidual(b))/(2*delta)
            return matrix
        def TestPointTrace(x):
            """Сохраняет читающий вывод исходного материала; координаты и все значения остаются точными Double."""
            nonlocal point_calls
            argument=win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8,tuple(float(z) for z in x))
            values=np.asarray(excel.Run(prefix+'TestReadExternalPointTrace',argument),float)
            assert values.ndim==2 and values.shape[1]==30 and np.all(np.isfinite(values))
            point_calls+=1
            actual_x=x
            if material is not None and material.get('mode')=='projected':
                actual_x=np.asarray(excel.Run(prefix+'TestMapExternalProjectedControl',argument),float).ravel()
            raw=struct.pack('<iii',len(x),*values.shape)+np.asarray(actual_x,dtype='<f8').tobytes()+values.astype('<f8').tobytes()
            (root/('Native/ExternalPoints.%06d.bin'%point_calls)).write_bytes(raw)
            return values
        initial=np.zeros(n)
        if (root/'Input/InitialCoordinates.bin').exists():
            raw=(root/'Input/InitialCoordinates.bin').read_bytes();count,=struct.unpack_from('<i',raw)
            assert count==n and len(raw)==4+8*n
            initial=np.frombuffer(raw,dtype='<f8',count=n,offset=4).copy()
        if material is not None and material.get('mode')=='projected':
            argument=win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8,tuple(float(z) for z in initial))
            before_projection=np.asarray(excel.Run(prefix+'TestEvaluateExternalEnvelope',argument),float).ravel()
            if not (root/'Input/InitialProjectedResidual.bin').exists() and not (root/'Input/ChangedInitialCondition.json').exists():
                expected=np.fromfile(root/'Input/InitialResidual.bin',dtype='<f8')
                assert np.array_equal(before_projection,expected),'Original native equations changed before enabling projection'
                result['original_native_equations_bits_equal_before_projection']=True
            (root/'Native/ExternalBeforeProjection.bin').write_bytes(before_projection.astype('<f8').tobytes())
            origin=initial
            if (root/'Input/ProjectionOrigin.bin').exists():
                raw=(root/'Input/ProjectionOrigin.bin').read_bytes();size,=struct.unpack_from('<i',raw)
                assert size==n and len(raw)==4+8*n
                origin=np.frombuffer(raw,dtype='<f8',count=n,offset=4)
            origin_argument=win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8,tuple(float(z) for z in origin))
            result['native_projection_metadata']=np.asarray(excel.Run(prefix+'TestOpenExternalProjectedControl',
                int(material['point']),int(material['side']),float(material['target_MPa']),origin_argument),float).ravel().tolist()
            result['residual_last_row_units']='normalized eliminated coordinate; original force/gauge rows remain MPa'
        first=TestResidual(initial); second=TestResidual(initial)
        assert np.array_equal(first,second),'Repeated immutable VBA evaluation differs'
        if material is not None and material.get('mode')=='projected' and (root/'Input/InitialProjectedResidual.bin').exists():
            expected=np.fromfile(root/'Input/InitialProjectedResidual.bin',dtype='<f8')
            assert np.array_equal(first,expected),'Cold identical projected native response differs'
            result['prior_projected_native_F_bits_equal']=True
        if (root/'Input/InitialResidual.bin').exists():
            expected=np.fromfile(root/'Input/InitialResidual.bin',dtype='<f8')
            if material is None:
                assert np.array_equal(first,expected),'Cold replay of the identical prior VBA residual differs'
                result['prior_context_residual_reproduced_bits_equal']=True
            elif material.get('mode')!='projected':
                assert np.array_equal(first[:-1],expected[:-1]),'Original native force/gauge rows changed'
                result['prior_original_native_force_gauge_rows_bits_equal']=True
                result['scalar_control_row_replaced_with_actual_native_trial_yield']=True
        solver_started=time.perf_counter()
        if method == 'diagnose':
            study=[]; previous_matrix=None; best=initial.copy(); best_norm=float(np.max(np.abs(first)))
            for relative in (1e-4,1e-5,1e-6,1e-7,1e-8,1e-9):
                matrix=np.empty((n,n));asymmetry=np.empty(n)
                for i in range(n):
                    delta=relative*(1+abs(initial[i]));a=initial.copy();b=initial.copy();a[i]+=delta;b[i]-=delta
                    plus=TestResidual(a);minus=TestResidual(b)
                    matrix[:,i]=(plus-minus)/(2*delta)
                    asymmetry[i]=np.max(np.abs(plus+minus-2*first))
                singular=np.linalg.svd(matrix,compute_uv=False)
                correction=np.linalg.lstsq(matrix,-first,rcond=None)[0]
                entry=dict(relative_step=relative,singular_values=singular.tolist(),
                    maximum_stencil_asymmetry_MPa=float(max(asymmetry)),
                    relative_matrix_change_from_previous=(float(np.linalg.norm(matrix-previous_matrix)/np.linalg.norm(matrix)) if previous_matrix is not None else None),
                    correction_coordinates=correction.tolist(),line_trials=[])
                for halving in range(11):
                    candidate=initial+correction*2.**(-halving)
                    value=TestResidual(candidate);norm=float(np.max(np.abs(value)))
                    entry['line_trials'].append(dict(halving=halving,residual_max_MPa=norm))
                    if norm<best_norm:best=candidate.copy();best_norm=norm
                study.append(entry);previous_matrix=matrix
                TestWrite(root/'Native/ExternalDerivativeStudy.json',dict(study=study,
                    original_force_tolerance_MPa=float(meta[3]),physical_failure=False,
                    exact_initial_residual_unchanged=bool(np.array_equal(TestResidual(initial),first))))
            answer=SimpleNamespace(x=best,success=False,message='Derivative sensitivity study; no iterative nonlinear solve',nfev=calls,njev=len(study))
            result['derivative_sensitivity_only']=True
        elif method == 'line':
            answer=run_path(str(root/'Study-TestExternalLine.py'))['TestStudy'](root,TestResidual,TestPointTrace,initial,first)
            result['readonly_material_line_study']=True
            result['point_trace_calls']=point_calls
        elif method == 'native':
            if material is None or material.get('mode')!='projected':raise ValueError('Native study requires the existing projected material adapter')
            argument=win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8,tuple(float(z) for z in initial))
            native_options=json.loads((root/'Input/NativeSolver.json').read_bytes()) if (root/'Input/NativeSolver.json').exists() else {}
            result['native_solver_options']=native_options
            native_solve=np.asarray(excel.Run(prefix+'TestSolveExternalProjectedNative',argument,float(step),bool(native_options.get('completeNewtonLineSearch',True))),float).ravel()
            assert len(native_solve)==n+7
            result['native_inner_solver_metadata']=native_solve[:7].tolist()
            result['native_inner_solver_is_SciPy']=False
            result['external_F_trace_excludes_inner_native_solver_calls']=True
            result['maximum_base_evaluations']=None
            answer=SimpleNamespace(x=native_solve[7:],success=bool(native_solve[0]==1),
                message='Existing VBA CCoupledHybridSolver; stop code '+str(int(native_solve[5])),
                nfev=int(native_solve[2]),njev=int(native_solve[3]))
        else:
            answer=least_squares(TestResidual,initial,jac=TestJacobian,method=method,x_scale='jac',
                ftol=1e-13,xtol=1e-13,gtol=1e-13,max_nfev=max_nfev,loss='linear',**({'tr_solver':'exact'} if method=='trf' else {}))
        final=TestResidual(answer.x);matrix=TestJacobian(answer.x)
        result.update(optimizer_success=bool(answer.success),optimizer_message=str(answer.message),base_evaluations=int(answer.nfev),
            optimizer_jacobians=int(answer.njev),actual_jacobians=jacobians,actual_VBA_function_calls=calls,
            solver_seconds=time.perf_counter()-solver_started,coordinates=answer.x.tolist(),residual_MPa=final.tolist(),
            residual_max_MPa=float(np.max(np.abs(final))),jacobian=matrix.tolist(),repeat_immutable_response_bits_equal=True)
        if method=='native':result['total_native_equation_function_calls']=calls+int(result['native_inner_solver_metadata'][2])
        argument=win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8,tuple(float(z) for z in answer.x))
        if material is not None and material.get('mode')=='projected':
            points=TestPointTrace(answer.x)
            selected=points[(points[:,0]==material['point'])&(points[:,1]==material['side'])]
            assert len(selected)==1
            result['final_native_material_constraint_MPa']=float(selected[0,17]-material['target_MPa'])
            result['force_gauge_residual_max_MPa']=float(max(abs(final[:-1])))
        native=np.asarray(excel.Run(prefix+'TestFinishExternalEnvelope',argument),float).ravel()
        assert len(native)==12+n-1
        binary=native.astype('<f8').tobytes();(root/'Native/ExternalNativeResult.bin').write_bytes(binary)
        result['native_result_binary_sha256']=sha256(binary).hexdigest().upper()
        result['native_verification']=native[:12].tolist();result['field']=native[12:].tolist()
        read=run_path(str(root/'Check-TestActualReverseCapture.py'))['TestReadBinary'];data=read(root/'Input/Reverse1.bin')
        operators=run_path(str(root/'Check-TestHardHalfSteps.py'))['TestOperators']
        import csv
        with (root/'Native/EnvelopeEnergy.Accepted.csv').open(encoding='cp1251') as f:
            dims=next(r[1:] for r in csv.reader(f,delimiter=';') if r and r[0]=='DIM')
        dofs,physical,points,concrete=map(int,dims[:4]);half=float(dims[4])
        g,load,w,v,coordinates,axis,b,slope=operators(root/'Native/EnvelopeEnergy.Operators.csv',dofs,points)
        def TestTensor(field):return np.asarray([b@field+s*half*(slope@field) for s in (-1,0,1)])
        def TestSquare(a):return float(np.einsum('spc,spc,p,c->',a,a,coordinates[:,2],np.array((1,1,1,.5,.5,.5)))/(3*sum(coordinates[:,2])))
        old=np.sqrt(TestSquare(TestTensor(data['fields'][-1]-data['fields'][-2])))
        distance=np.sqrt(TestSquare(TestTensor(native[12:]-data['fields'][-1])))
        predicted=abs(meta[7])*old
        result['tensor_distance']=float(distance);result['predicted_tensor_distance']=float(predicted)
        local=bool(distance<=3*predicted and predicted>0)
        result['within_three_predicted_distances']=local
        verified=bool(native[0]==1 and native[1]==1 and native[2]==1 and native[3]<=1 and native[4]>0 and
            np.max(np.abs(final))<=native[10] and native[5]-meta[1]>2*native[10]/meta[4] and local)
        result['nds_verified_on_local_forward_energy_step']=verified
        if material is not None:
            result['nds_verified_on_local_forward_energy_step']=False
            result['native_original_force_NDS_verified']=bool(native[0]==1 and native[1]==1)
            result['native_material_coordinate_verified']=bool(abs(result['final_native_material_constraint_MPa'])<=2e-12) if material.get('mode')=='projected' else bool(abs(final[-1])<=native[10])
            result['material_coordinate_branch_connection_proven']=False
            result['material_coordinate_forward_tensor']=bool(native[2]==1) if native[0]==1 else None
            result['material_target_not_a_strength_criterion']=True
        try:
            weights=np.linalg.solve(matrix.T,np.eye(n)[:,-1]);condition=float(np.linalg.cond(matrix))
            result['jacobian_condition']=condition if np.isfinite(condition) else None
            result['jacobian_condition_nonfinite']=not bool(np.isfinite(condition))
            result['diagnostic_lambda_uncertainty']=float(abs(scales[-1])*max(native[10],np.max(np.abs(final)))*sum(abs(weights))) if condition<1e14 else None
        except np.linalg.LinAlgError:result['diagnostic_lambda_uncertainty']=None
        result['outcome']='VerifiedLocalNDS;NotCapacity' if verified else 'NoVerifiedLocalNDS;NotPhysicalFailure'
        if material is not None:
            valid=bool(native[0]==1 and native[1]==1 and native[3]<=1 and native[4]>0 and
                max(abs(final[:-1]))<=native[10] and result['native_material_coordinate_verified'])
            if material.get('mode')=='projected':
                valid=bool(valid and abs(final[-1])<=1e-10)
                result['diagnostic_lambda_uncertainty']=None
                result['diagnostic_lambda_uncertainty_unavailable_reason']='row-specific bounds required for dimensionless pivot equation'
            result['outcome']='VerifiedNDSAtMaterialCoordinate;NotCapacity;BranchConnectionUnproven' if valid else 'NoVerifiedMaterialCoordinateNDS;NotPhysicalFailure'
            result['point_trace_calls']=point_calls
    except Exception as error:
        result.update(outcome='ExternalControlError;NotPhysicalFailure',error=str(error),traceback=traceback.format_exc())
    finally:
        result['elapsed_seconds']=time.perf_counter()-started
        TestWrite(root/'Native/ExternalNumericalResult.json',result)
        if trace is not None:trace.close()
        if owner:
            if book is not None:
                try:
                    if Path(str(book.FullName)).resolve()!=fixture.resolve():raise RuntimeError('Owned fixture identity changed')
                    book.Close(False)
                except Exception as error:result['book_cleanup_error']=str(error)
                book=None
            if excel is not None:
                try:
                    if excel.Workbooks.Count:
                        owner['protected_workbooks_detected']=True
                        result['cleanup_protected_unknown_workbooks']=True
                    else:excel.Quit()
                except Exception as error:result['excel_quit_error']=str(error)
        excel=None;gc.collect();pythoncom.CoUninitialize()
        if owner:
            for _ in range(30):
                state=TestProcess(owner['pid'],owner['filetime'])
                if state['state']!='Alive':break
                time.sleep(.2)
            if state['state']=='Alive' and not owner.get('protected_workbooks_detected'):
                result['owned_empty_excel_cleanup_termination']=TestProcess(owner['pid'],owner['filetime'],True)
                time.sleep(.5);state=TestProcess(owner['pid'],owner['filetime'])
            result['owned_excel_final_state']=state
        result['source_unchanged']=sha256(source.read_bytes()).hexdigest().upper()==source_hash
        TestWrite(output,result);done.set()
        if timer is not None:timer.cancel()
    print('EXTERNAL EXACT VBA CONTROL:',result['outcome'],'calls=',result.get('actual_VBA_function_calls'),flush=True)
    return 1 if result['outcome'].startswith('ExternalControlError') or not result['source_unchanged'] or result.get('owned_excel_final_state',{}).get('state') not in ('Absent','Replaced','Exited') else 0


if __name__=='__main__':
    p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,required=True);p.add_argument('--method',choices=('trf','lm','diagnose','line','native'),default='trf')
    p.add_argument('--seed-kind',type=int,choices=(0,1),default=1);p.add_argument('--fraction',type=float,default=.000015625)
    p.add_argument('--step',type=float,default=1e-6);p.add_argument('--max-nfev',type=int,default=40);p.add_argument('--timeout',type=int,default=600)
    a=p.parse_args();raise SystemExit(TestRun(a.root.resolve(),a.method,a.seed_kind,a.fraction,a.step,a.max_nfev,a.timeout))
