"""ДЛЯ ТЕСТОВ. Проверяет локальные численные остановки, три равновесия промежуточных целей и точные уравнения внешнего контроля."""
from argparse import ArgumentParser
from contextlib import redirect_stdout
from hashlib import sha256
from io import StringIO
import csv,io,json,re,struct,tarfile
from pathlib import Path
from runpy import run_path
from tempfile import TemporaryDirectory
import numpy as np


def TestExternal(folder):
    """Отделяет остановку/успех оптимизатора от исходного равновесия; проверяет весь принятый префикс независимо."""
    here=Path(__file__).resolve().parent
    result=json.loads((folder/'Native/ExternalResult.json').read_bytes())
    owner=json.loads((folder/'Native/ExternalOwnership.json').read_bytes())
    assert result['source_unchanged'] and result['owned_excel_final_state']['state']=='Exited'
    assert result['owned_excel_final_state']['filetime']==owner['filetime']
    assert result['working_external_dependency_added'] is False and result['physical_capacity_certified'] is False
    n=61;raw=(folder/'Native/ExternalFTrace.bin').read_bytes()
    assert raw[:16]==b'NDMEXTFTRACE0001' and struct.unpack_from('<i',raw,16)[0]==n
    dtype=np.dtype([('seq','<u8'),('seconds','<f8'),('x','<f8',(n,)),('residual','<f8',(n,))])
    records=np.frombuffer(raw,dtype=dtype,offset=20)
    assert len(records)==result['actual_VBA_function_calls']
    assert np.array_equal(records['seq'],np.arange(1,len(records)+1))
    assert np.all(np.isfinite(records['x'])) and np.all(np.isfinite(records['residual']))
    assert np.array_equal(records['x'][0],records['x'][1]) and np.array_equal(records['residual'][0],records['residual'][1])
    coordinates=np.asarray(result['coordinates'],float);residual=np.asarray(result['residual_MPa'],float)
    hit=np.all(records['x']==coordinates,axis=1)
    assert np.any(hit) and all(np.array_equal(v,residual) for v in records['residual'][hit])
    assert max(abs(residual))==result['residual_max_MPa']
    native=np.frombuffer((folder/'Native/ExternalNativeResult.bin').read_bytes(),dtype='<f8')
    assert native.shape==(72,) and native.tolist()==result['native_verification']+result['field']
    assert native[11]==result['context_metadata'][10]
    seed_raw=(folder/'Input/ExternalSeed.bin').read_bytes()
    assert struct.unpack_from('<i',seed_raw)[0]==61 and len(seed_raw)==492
    seed=np.frombuffer(seed_raw,dtype='<f8',offset=4)
    mapped=seed+coordinates*np.asarray(result['unknown_scales'],float)
    np.testing.assert_allclose(mapped[:-1],native[12:],rtol=1e-14,atol=1e-300)
    assert mapped[-1]==native[4]
    target=struct.unpack('<id',(folder/'Input/EnvelopeFailureLevel.bin').read_bytes())[1]
    assert abs((native[5]-target)*result['context_metadata'][4]-residual[-1])<1e-12
    helpers=run_path(str(here/'Validate-TestHardPrepared.py'))
    accepted,fields=helpers['TestValidate'](folder,'EnvelopeEnergy.Accepted')
    original=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](folder/'Input/Reverse1.bin')
    np.testing.assert_allclose(fields,original['fields'],rtol=1e-14,atol=1e-300)
    return dict(outcome=result['outcome'],optimizer_success=result['optimizer_success'],
        optimizer_message=result['optimizer_message'],base_evaluations=result['base_evaluations'],
        calls=len(records),residual_max_MPa=float(max(abs(residual))),original_tolerance_MPa=result['context_metadata'][3],
        native_force_verified=bool(native[0]),accepted_revision_unchanged=True,
        duplicate_native_response_bits_equal=True,accepted_history_independent_equilibrium=accepted,
        own_excel_closed=True,physical_capacity_certified=False)


def TestProof(repo,root):
    """Привязывает результаты к фактическому VBA, исходным данным и сохранённым журналам, не повышая численную остановку до физического отказа."""
    here=Path(__file__).resolve().parent
    for name,digest in json.loads((root/'Files.json').read_bytes()).items():
        assert sha256((root/name).read_bytes()).hexdigest().upper()==digest,name
    class TestRepo:
        """Использует сохранённую базовую книгу и закреплённые декодеры проекта."""
        def __truediv__(self,name):
            return root/'BaseWorkbook.xlsm' if str(name)=='workbook/output/RC_Section_NDM.xlsm' else repo/name
    base=TestRepo();bundle=json.loads((root/'CompiledProjects.json').read_bytes());binaries={}
    with tarfile.open(root/'CompiledProjects.tar.xz',mode='r|xz') as archive:
        for member in archive:
            assert member.isfile() and '/' not in member.name and member.size<=64*1024*1024
            key=member.name[:-4];raw=archive.extractfile(member).read()
            assert sha256(raw).hexdigest().upper()==bundle['projects'][key]['sha256'];binaries[key]=raw
    helper,_,reader=run_path(str(here/'Inspect-TestResolutionVBA.py'))['TestHelpers'](base)
    decoded={}
    def TestComponents(key):
        """Кэширует только точные неизменяемые CFB-байты в рамках одной проверки."""
        digest=sha256(binaries[key]).hexdigest().upper()
        if digest not in decoded:_,decoded[digest]=reader(helper['TestContainer'](binaries[key]))
        return decoded[digest]
    def TestIdentifierCase(code):
        """Учитывает только регистр идентификаторов VBE, сохраняя все строки и комментарии."""
        return re.sub(r'"(?:[^"\n]|"")*"|\'[^\n]*|[A-Za-z_][\w]*',lambda m:m.group() if m.group().startswith(('"',"'")) else m.group().lower(),code)
    source=(here/'Audit-TestPrivateCoupledBook.py').read_text('utf-8-sig')
    anchor="    assert compilation['Compiled'] and not compilation['ProductionChanged']";assert source.count(anchor)==1
    source=source.replace(anchor,anchor+"\n    compilation['Modules']=json.loads((root/'Preparation.json').read_bytes())['modules']",1)
    namespace={'__name__':'failure_source_audit_test','__file__':str(here/'Audit-TestPrivateCoupledBook.py')}
    exec(compile(source,namespace['__file__'],'exec'),namespace)
    cases=json.loads((root/'Cases.json').read_bytes());parents=json.loads((root/'ParentProjects.json').read_bytes());sources={}
    expected={'Precision':['CCoupledConcreteResponse','CCoupledConcreteReturn','modTestCoupledElasticState'],
        'Following':['CCoupledDeformationContinuation'],'OneSided':['CCoupledHybridSolver','modTestCoupledElasticState'],'TRFRestart':[]}
    for key in cases:
        folder=root/'Cases'/key;prep=json.loads((folder/'Preparation.json').read_bytes())
        execution=json.loads((folder/'CaseExecution.json').read_bytes());imports=execution['actual_imports']
        assert imports==expected.get(key,['modTestCoupledElasticState']),(key,imports)
        with TemporaryDirectory() as temporary,redirect_stdout(StringIO()):
            out=Path(temporary)/'Source.json';namespace['TestAudit'](base,folder,out);audit=json.loads(out.read_bytes())
        assert (audit['components'],audit['imported'],audit['unchanged'])==(254,41,213)
        parent,actual=TestComponents(parents[key]['case']),TestComponents(key)
        assert sha256(binaries[parents[key]['case']]).hexdigest().upper()==parents[key]['sha256']
        assert len(parent)==len(actual)==254 and set(parent)==set(actual)
        assert all(TestIdentifierCase(parent[n])==TestIdentifierCase(actual[n]) for n in actual if n not in imports)
        if not imports:assert binaries[key]==binaries[parents[key]['case']]
        for file in (root/'WorkingSources').iterdir():
            candidate=root/'ExperimentalSources'/key/file.name
            category='Materials' if file.name in ('CCoupledConcreteResponse.cls','CCoupledConcreteReturn.cls','CCoupledReinforcedResponse.cls') else 'Solver'
            prepared=(folder/'src'/category/file.name).read_text('utf-8-sig')
            if candidate.exists():assert prepared==candidate.read_text('utf-8-sig')
            else:
                working=file.read_text('utf-8-sig')
                if file.name=='CCoupledHybridSolver.cls':
                    line='                modTestCoupledElasticState.TestRecordLiveDifference before, after, score, j, mJacobianCalls, mSteps, mFunctionCalls\n'
                    assert prepared.count(line)==1;prepared=prepared.replace(line,'',1)
                assert prepared.startswith(working),(key,file.name)
                assert not prepared[len(working):] or 'ДЛЯ ТЕСТОВ' in prepared[len(working):]
        sources[key]=dict(source=audit,actual_imports=imports,other_components_exact_except_identifier_case=254-len(imports))
    numerical={key:run_path(str(here/'Check-TestEnvelopeDifferenceBatch.py'))['TestCheck'](root/'Cases'/key)
        for key in ('Difference','Precision','Actual','Homotopy','OneSided')}
    for key in ('Difference','Precision','Actual','OneSided'):assert not any(v['strict_solver_success'] for v in numerical[key]['cases'].values())
    assert numerical['Homotopy']['exported_original_force_candidates']==[1,2,3]
    assert all(numerical['Homotopy']['cases'][i]['strict_solver_success']==(i<=3) for i in range(1,6))
    comparison=run_path(str(here/'Compare-TestEnvelopeMaterialPrecision.py'))['TestCompare'](root/'Cases/Difference',root/'Cases/Precision')
    assert json.loads(json.dumps(comparison))==json.loads((root/'MaterialPrecisionComparison.json').read_bytes())
    following=root/'Cases/Following'
    timeout=json.loads((following/'Native/ExternalTimeout.json').read_bytes());closed=json.loads((following/'Native/TimeoutClosure.json').read_bytes())
    assert closed['after']['state'] in ('Absent','Exited') and closed['owner']['filetime']==timeout['owner']['filetime']
    assert timeout['timeout_seconds']==600 and not timeout['physical_failure']
    assert not json.loads((following/'PolicyDecision.json').read_bytes())['retained']
    checkpoint=(following/'Native/Controller.Checkpoints.csv').read_bytes()
    origin=json.loads((root/'Cases/Actual/Input/FailureOrigin.json').read_bytes())
    assert sha256(checkpoint).hexdigest().upper()==origin['source_checkpoint_sha256']
    assert origin['count']==200 and origin['before_call']==66 and origin['after_call']==67
    blocks={};current=None
    for row in csv.reader(io.StringIO(checkpoint.decode('cp1251')),delimiter=';'):
        if not row:continue
        if row[0]=='CHECKPOINT':current=dict(call=int(row[1]),count=int(row[2]),points=[],details='')
        elif current is not None and row[0]=='DETAILS':current['details']=';'.join(row[1:])
        elif current is not None and row[0]=='P':current['points'].append(row[1:])
        elif current is not None and row[0]=='CHECKPOINT_END':
            assert int(row[1])==current['call'] and len(current['points'])==current['count']
            blocks[current['call']]=current;current=None
    assert len(blocks)==77 and blocks[max(blocks)]['count']==206
    first,second=blocks[66],blocks[67]
    assert first['points']==second['points'] and len(second['points'])==200
    assert first['details']!=second['details']
    assert 'остановка 5' in second['details']
    read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary']
    accepted=read(root/'Cases/Actual/Input/Reverse1.bin');initial=read(following/'Input/Reverse1.bin')
    assert len(initial['fields'])==147 and np.array_equal(accepted['fields'][:147],initial['fields'])
    decimal=np.asarray([r[3:] for r in second['points']],float)
    assert decimal.shape==(200,60) and np.array_equal(accepted['fields'][147:],decimal[147:])
    np.testing.assert_allclose(accepted['fields'][:147],decimal[:147],rtol=1e-14,atol=1e-300)
    assert abs(float(second['points'][-1][2])-1.11067846092076)<1e-14
    pair=b''
    for number in (66,67):
        begin=checkpoint.index(('CHECKPOINT;'+str(number)+';').encode())
        end=checkpoint.index(('CHECKPOINT_END;'+str(number)).encode(),begin)
        end=checkpoint.find(b'\n',end)+1;assert end>begin;pair+=checkpoint[begin:end]
    assert pair==(root/'Cases/Actual/Input/FailureCheckpointPair.csv').read_bytes()
    external={key:TestExternal(root/'Cases'/key) for key in ('TRF','TRFRestart')}
    first,second=root/'Cases/TRF',root/'Cases/TRFRestart'
    assert (first/'Input/Reverse1.bin').read_bytes()==(second/'Input/Reverse1.bin').read_bytes()
    raw=(first/'Native/ExternalNativeResult.bin').read_bytes();values=struct.unpack('<72d',raw)
    assert (second/'Input/ExternalSeed.bin').read_bytes()==struct.pack('<i61d',61,*values[12:],values[4])
    assert json.loads((root/'SetupFailure/CompilePointClosure.json').read_bytes())['State']=='Absent'
    return dict(source=sources,numerical=numerical,local_material_precision_comparison=comparison['comparisons'],
        following_policy=dict(outcome='Timeout600',retained=False,own_excel_closed=True,independent_equilibrium_of_its_entire_path_not_claimed=True),
        external=external,external_restart_seed_exact_Double=True,working_sources_unchanged_from_checkpoint637=True,
        installed_workbook_changed=False,physical_capacity_certified=False,full_acceptance_completed=False,
        independent_material_reintegration=False,initial_test_compile_failure_preserved=True)


if __name__=='__main__':
    p=ArgumentParser(description=__doc__);p.add_argument('--repo',type=Path,default=Path.cwd());p.add_argument('--root',type=Path,required=True);p.add_argument('--out',type=Path,required=True)
    a=p.parse_args();proof=TestProof(a.repo.resolve(),a.root.resolve());a.out.write_bytes((json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode())
    print('ENVELOPE FAILURE EVIDENCE VERIFIED; THREE INTERMEDIATE EQUILIBRIA; PHYSICAL CAPACITY OPEN')
