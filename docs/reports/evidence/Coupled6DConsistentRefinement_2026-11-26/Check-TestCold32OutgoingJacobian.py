"""ДЛЯ ТЕСТОВ. Проверяет фактический рабочий прогноз по матрице уровня1/1024 и сохраняет его направленную согласованность."""
from pathlib import Path
here=Path(__file__).resolve().parent;file=here/'Check-TestCold32DirectionalJacobian.py'
code=file.read_text('utf-8-sig').replace('Cold899Cold32DirectionalJacobian_20261126','Cold900Cold32OutgoingJacobian_20261126')
code=code.replace('for level in (0,1/64,1,8):','for level in (0,1/1024,1/64,1,8):')
anchor='meta=r.VectorTest();accepted=r.VectorTest();movement=r.VectorTest();projection=r.VectorTest()\n'
assert code.count(anchor)==1
code=code.replace(anchor,anchor+'''outgoing_ready=r.VectorTest();assert outgoing_ready.tolist()==[1.]
outgoing=r.VectorTest();outgoing_probe=r.VectorTest()
np.testing.assert_array_equal(outgoing_probe,accepted+movement/1024)
outgoing_check={}
''',1)
anchor="    predicted={'refined':refined@direction,'ordinary':ordinary@direction}\n"
assert code.count(anchor)==1
code=code.replace(anchor,anchor+'''    if level==1/1024:
        rhs=-refined[:60,60]/scales[60]
        independent=np.linalg.solve(refined[:60,:60],rhs)*scales[:60]
        relative=float(np.max(abs(independent-outgoing))/max(np.max(abs(outgoing)),1e-300))
        defect=refined[:60,:60]@(outgoing/scales[:60])-rhs
        assert relative<1e-8
        op=np.array([B+s*.001*gradient for s in (-1,0,1)])
        M=np.einsum('spcd,spce,p,c->de',op,op,coords[:,2],np.array([1,1,1,.5,.5,.5]))/(3*sum(coords[:,2]));M[56:]=0;M[:,56:]=0
        cosine=float(outgoing@M@movement/np.sqrt((outgoing@M@outgoing)*(movement@M@movement)))
        outgoing_check.update(relative_difference_from_independent_linear_solve=relative,
            normalized_linear_residual_MPa=float(max(abs(defect))),cosine_with_incoming_secant=cosine,
            native_TrySpatialLoadTangent_probe_is_exact_1_over1024=True,predictor_not_equilibrium=True)
''',1)
anchor=' derivative_discrepancies_preserved=True,levels=results,'
assert code.count(anchor)==1
code=code.replace(anchor,' outgoing_working_predictor=outgoing_check,derivative_discrepancies_preserved=True,levels=results,',1)
code=code.replace("root/'Cold32DirectionalJacobianProof.json'","root/'Cold32OutgoingJacobianProof.json'")
code+='\nprint("OUTGOING_WORKING_PREDICTOR",outgoing_check)\n'
exec(compile(code,str(file),'exec'),{'__file__':str(file),'__name__':'cold32_outgoing_jacobian_test'})
