"""ДЛЯ ТЕСТОВ. Сравнивает штатный Newton и SciPy TRF на одинаковых60 уравнениях при фиксированных шести усилиях; статус оптимизатора не заменяет равновесие."""
from argparse import ArgumentParser
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from uuid import uuid4
import json,shutil,threading,time,os,gc,struct,traceback,ctypes
import numpy as np
import scipy
from scipy.optimize import least_squares,root as scipy_root

here=Path(__file__).resolve().parent
shared=run_path(str(here/'Run-TestExternalEnvelope.py'))
pythoncom=shared['pythoncom'];win32com=shared['win32com'];win32process=shared['win32process']
TestProcess=shared['TestProcess'];TestWrite=shared['TestWrite']

def TestRun(root,maximum,timeout,method):
    """Внешний алгоритм получает только невязки; материал и история совпадают со штатной силовой пробой внутри того же собственного Excel."""
    prep=json.loads((root/'Preparation.json').read_bytes())
    for name,digest in prep['inputs'].items():assert sha256((root/name).read_bytes()).hexdigest().upper()==digest,name
    compiled=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
    source=root/'docs/regression/Diagnostic.xlsm';source_hash=sha256(source.read_bytes()).hexdigest().upper()
    assert source_hash==compiled['WorkbookSHA256'] and compiled['Compiled'] and not compiled['ProductionChanged']
    assert json.loads((root/'CompilePointClosure.json').read_bytes())['State']=='Absent'
    output=root/'Native/FixedForceExternalResult.json';assert not output.exists()
    fixture=root/'Temp'/('FixedForce_'+uuid4().hex)/'Diagnostic.xlsm';fixture.parent.mkdir();shutil.copyfile(source,fixture)
    result=dict(source_sha256=source_hash,scipy_version=scipy.__version__,numpy_version=np.__version__,method=method,maximum_base_evaluations=maximum,
        same_VBA60_force_equations=True,external_material_implementation=False,fixed_lambda_not_unknown=True,
        physical_capacity_certified=False,working_external_dependency_added=False)
    before=set(win32process.EnumProcesses());owner={};excel=None;book=None;stream=None;done=threading.Event();timer=None
    started=time.perf_counter();calls=0;jacobians=0;last=None
    def TestTimeout():
        """Прерывает только доказанно собственный Excel с неизменной личностью; таймаут не является физическим отказом."""
        if done.is_set():return
        TestWrite(root/'Native/FixedForceTimeout.json',dict(owner=owner,timeout=timeout,physical_failure=False))
        if owner.get('exclusive') and not owner.get('protected_workbooks'):
            TestProcess(owner['pid'],owner['filetime'],True)
        os._exit(124)
    pythoncom.CoInitialize()
    try:
        excel=win32com.client.dynamic.Dispatch(pythoncom.CoCreateInstance('Excel.Application',None,pythoncom.CLSCTX_LOCAL_SERVER,pythoncom.IID_IDispatch))
        _,pid=win32process.GetWindowThreadProcessId(int(excel.Hwnd));assert pid not in before,'Existing Excel must not be used'
        identity=TestProcess(pid);assert identity['state']=='Alive'
        owner.update(pid=pid,filetime=identity['filetime'],start_ticks_utc=str(identity['start_ticks_utc']))
        TestWrite(root/'Native/FixedForceOwner.json',owner)
        timer=threading.Timer(timeout,TestTimeout);timer.daemon=True;timer.start()
        excel.Visible=False;excel.DisplayAlerts=False;excel.EnableEvents=False;excel.ScreenUpdating=False;excel.AutomationSecurity=1
        assert excel.Workbooks.Count==0
        book=excel.Workbooks.Open(str(fixture),0,False);assert not book.ReadOnly and Path(str(book.FullName)).resolve()==fixture.resolve()
        owner['exclusive']=True;TestWrite(root/'Native/FixedForceOwner.json',owner)
        prefix="'"+book.Name.replace("'","''")+"'!modTestCoupledElasticState."
        meta=np.asarray(excel.Run(prefix+'TestStartFixedForceComparison',str(root)),float).ravel();assert meta.shape==(5,) and meta[0]==60
        result['metadata']=meta.tolist();result['native_probe_finished_before_external_solver']=True
        scales=np.asarray(excel.Run(prefix+'TestExternalEnvelopeScales'),float).ravel();assert scales.shape==(61,)
        result['coordinate_scales']=scales.tolist()
        stream=(root/'Native/FixedForceFTrace.bin').open('xb');stream.write(b'NDMFIXFTRACE0001'+struct.pack('<i',60));stream.flush()
        def TestResidual(coordinates):
            """Записывает фактически вычисленные Double; не подставляет штрафы вместо отказавшего отклика материала."""
            nonlocal calls,last
            coordinates=np.asarray(coordinates,float);assert coordinates.shape==(60,) and np.all(np.isfinite(coordinates))
            last=coordinates.copy()
            argument=win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8,tuple(float(x) for x in coordinates))
            values=np.asarray(excel.Run(prefix+'TestEvaluateFixedForceComparison',argument),float).ravel();assert values.shape==(60,) and np.all(np.isfinite(values))
            calls+=1;stream.write(struct.pack('<Qd',calls,time.perf_counter()-started));stream.write(coordinates.astype('<f8').tobytes());stream.write(values.astype('<f8').tobytes());stream.flush()
            if calls%500==0:print('FIXED_FORCE_CALLS',calls,'MAX_MPa',float(max(abs(values))),flush=True)
            return values
        def TestJacobian(coordinates):
            """Строит центральную разность тех же60 уравнений, не меняя lambda или материальную историю."""
            nonlocal jacobians
            jacobians+=1;matrix=np.empty((60,60))
            for i in range(60):
                step=1e-6*(1+abs(coordinates[i]));a=coordinates.copy();b=a.copy();a[i]+=step;b[i]-=step
                matrix[:,i]=(TestResidual(a)-TestResidual(b))/(a[i]-b[i])
            return matrix
        initial=np.zeros(60);first=TestResidual(initial);second=TestResidual(initial);assert np.array_equal(first,second)
        result['initial_residual_max_MPa']=float(max(abs(first)));result['repeat_response_bitexact']=True
        solve_started=time.perf_counter()
        if max(abs(first))<=meta[4]:
            answer=initial;result.update(optimizer_invoked=False,optimizer_success=None,optimizer_message='Initial field already meets weak-equation tolerance')
        else:
            if method=='hybr':
                solved=scipy_root(TestResidual,initial,jac=TestJacobian,method='hybr',options=dict(xtol=1e-10,maxfev=maximum,factor=.2))
            else:
                solved=least_squares(TestResidual,initial,jac=TestJacobian,method='trf',x_scale='jac',tr_solver='exact',loss='linear',
                    ftol=1e-13,xtol=1e-13,gtol=1e-13,max_nfev=maximum)
            answer=solved.x;result.update(optimizer_invoked=True,optimizer_success=bool(solved.success),optimizer_message=str(solved.message),
                base_evaluations=int(solved.nfev),optimizer_jacobians=int(solved.njev))
        result['solver_seconds']=time.perf_counter()-solve_started
        final=TestResidual(answer)
        argument=win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8,tuple(float(x) for x in answer))
        checked=np.asarray(excel.Run(prefix+'TestFinishFixedForceComparison',argument),float).ravel();assert checked.shape==(5,) and checked[0]==meta[1] and checked[3]==checked[4]==meta[3]
        result.update(coordinates=answer.tolist(),final_residual=final.tolist(),final_max_MPa=float(max(abs(final))),native_final_metadata=checked.tolist(),
            externally_found_equilibrium=bool(checked[2]<=checked[1]),outcome='EquilibriumFound' if checked[2]<=checked[1] else 'NoEquilibriumFoundWithinBudget;NotPhysicalFailure')
        assert (root/'Native/FixedForceAcceptedBefore.bin').read_bytes()==(root/'Native/FixedForceAfterNative.bin').read_bytes()==(root/'Native/FixedForceAcceptedAfter.bin').read_bytes()
        result['accepted_history_unchanged_bitexact']=True
    except Exception as error:
        result.update(outcome='ExternalProbeError;NotPhysicalFailure',error=str(error),traceback=traceback.format_exc(),last_coordinates=None if last is None else last.tolist())
    finally:
        if stream is not None:stream.close()
        if owner:
            if book is not None:
                try:
                    assert Path(str(book.FullName)).resolve()==fixture.resolve();book.Close(False)
                except Exception as error:result['book_cleanup_error']=str(error)
                book=None
            if excel is not None:
                try:
                    if excel.Workbooks.Count:owner['protected_workbooks']=True
                    else:excel.Quit()
                except Exception as error:result['excel_cleanup_error']=str(error)
        excel=None;gc.collect();pythoncom.CoUninitialize()
        if owner:
            for _ in range(30):
                state=TestProcess(owner['pid'],owner['filetime'])
                if state['state']!='Alive':break
                time.sleep(.2)
            if state['state']=='Alive' and not owner.get('protected_workbooks'):
                result['owned_empty_cleanup']=TestProcess(owner['pid'],owner['filetime'],True);time.sleep(.5);state=TestProcess(owner['pid'],owner['filetime'])
            result['own_excel_final_state']=state
        result.update(actual_function_calls=calls,actual_jacobians=jacobians,elapsed_seconds=time.perf_counter()-started,
            source_unchanged=sha256(source.read_bytes()).hexdigest().upper()==source_hash)
        TestWrite(output,result);done.set()
        if timer is not None:timer.cancel()
    print('FIXED FORCE EXTERNAL',result['outcome'],'CALLS',calls,'MAX',result.get('native_final_metadata'),flush=True)
    return 1 if result['outcome'].startswith('ExternalProbeError') or not result['source_unchanged'] or result.get('own_excel_final_state',{}).get('state') not in ('Absent','Exited','Replaced') else 0

if __name__=='__main__':
    parser=ArgumentParser(description=__doc__);parser.add_argument('--root',type=Path,required=True);parser.add_argument('--max-nfev',type=int,default=30);parser.add_argument('--timeout',type=int,default=600)
    parser.add_argument('--method',choices=['trf','hybr'],default='trf')
    args=parser.parse_args();raise SystemExit(TestRun(args.root.resolve(),args.max_nfev,args.timeout,args.method))
