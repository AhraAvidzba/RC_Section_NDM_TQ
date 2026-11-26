"""ДЛЯ ТЕСТОВ: готовит частную книгу для проверки рабочего спектрального материала."""

from argparse import ArgumentParser
from hashlib import sha256
import json
import math
from pathlib import Path
from subprocess import check_output
from zipfile import ZipFile
from io import BytesIO


def TestStage(repo, root):
    """Закрепляет инструменты, рабочие исходники и две независимые калибровки."""
    here = Path(__file__).resolve().parent
    template = Path('D:/CodexScratch/RC_Section_NDM_TQ/ForceSubspaceStock20261118')
    prep = json.loads((template/'Preparation.json').read_bytes())
    files = {}
    for name, expected in prep['inputs'].items():
        if name.startswith('tools/') or name == 'Guard-IsolatedOffice.ps1':
            raw = (template/name).read_bytes()
            assert sha256(raw).hexdigest().upper() == expected
            files[name] = raw
    source = (repo/'workbook/output/RC_Section_NDM.xlsm').read_bytes()
    expected = 'E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
    assert sha256(source).hexdigest().upper() == expected
    files['workbook/output/RC_Section_NDM.xlsm'] = source
    files['docs/regression/Diagnostic.xlsm'] = source
    names = ['CCubicSplineBasis', 'CCoupledSectionQuadrature', 'CCoupledElasticState',
             'CCoupledFieldSubspace', 'CCoupledNewtonSolver', 'CCoupledConcreteBackbone',
             'CCoupledASDFlow', 'CCoupledASDTangent', 'CCoupledConcreteResponse', 'CCoupledReinforcedResponse', 'CCoupledMaterialSnapshot',
             'CCoupledSectionAnalysis', 'CNDMResultsWriter', 'modCoupledConcreteTensor', 'modTestCoupledASDMaterial', 'modTestCoupledASDTangent',
             'CCoupledDiscretization', 'modTestCoupledQuadrature', 'CCoupledMaterialFailure', 'CCoupledDeformationContinuation',
             'ICoupledContinuationProgress', 'CCoupledProgressObserverTest', 'CCoupledPrincipalTransition']
    for name in names:
        paths = list(repo.glob('src/**/'+name+'.cls'))+list(repo.glob('src/**/'+name+'.bas'))+list(repo.glob('tests/'+name+'.bas'))+list(repo.glob('tests/'+name+'.cls'))
        assert len(paths) == 1, name
        file = paths[0]
        raw = file.read_bytes()
        assert max(map(len, raw.decode('utf-8-sig').splitlines())) < 1024
        files[file.relative_to(repo).as_posix()] = raw
    for name in ('default', 'ndm_uniaxial'):
        files['Data/'+name+'.csv'] = (Path('D:/CodexScratch/RC_Section_NDM_TQ/ClosedASDG5S220261118')/(name+'.csv')).read_bytes()
    mvt=Path('D:/CodexScratch/RC_Section_NDM_TQ/ForceSubspaceMVT0220261118')
    mvt_prep=json.loads((mvt/'Preparation.json').read_bytes())
    relative='tests/modTestMVT2WorkingSubspaceTest.bas'
    raw=(mvt/relative).read_bytes()
    assert sha256(raw).hexdigest().upper()==mvt_prep['inputs'][relative]
    code=raw.decode('utf-8-sig').replace('\r\n','\n')
    code=code.replace('modTestMVT2WorkingSubspaceTest','modTestMVT2ProductionASDTest').replace('RunTestMVT2WorkingSubspaceTest','RunTestMVT2ProductionASDTest')
    code=code.replace('TOTAL_MVT2_WORKING_SUBSPACE_TEST','TOTAL_MVT2_PRODUCTION_ASD_TEST')
    declaration='    Dim pointCount As Long: pointCount = operators.PointCount'
    assert code.count(declaration)==1
    code=code.replace(declaration,declaration+'\n    Dim flow As CCoupledASDFlow\n    Set flow = modTestCoupledASDMaterial.TestReadFlow(Environ$("NDM_MVT_FLOW_FOLDER") & "\\Data", "default")',1)
    tail='pointCount, ec, nu, fc0, fc, ft, peakStrain, gt, gc, LOCALIZATION_LENGTH, dilation:=32#'
    assert code.count(tail)==3
    code=code.replace('concrete.Initialize '+tail,'concrete.InitializeASD pointCount, flow')
    code=code.replace('plusConcrete.Initialize '+tail,'plusConcrete.InitializeASD pointCount, flow')
    code=code.replace('minusConcrete.Initialize '+tail,'minusConcrete.InitializeASD pointCount, flow')
    old='    replay.Initialize operators.PointCount, modulus, poisson, 0.6 * fc, fc, ft, peak, gt, gc, 125#, dilation:=32#'
    assert code.count(old)==1
    code=code.replace(old,'    replay.InitializeASD operators.PointCount, modTestCoupledASDMaterial.TestReadFlow(Environ$("NDM_MVT_FLOW_FOLDER") & "\\Data", "default")')
    old='report & "TOTAL_MVT2_PRODUCTION_ASD_TEST: passed=1; failed=0; target_reached=" & CStr(completed)'
    assert code.count(old)==1
    code=code.replace(old,'report & "TOTAL_MVT2_PRODUCTION_ASD_TEST: " & IIf(completed, "passed=1; failed=0", "passed=0; failed=1") & "; target_reached=" & CStr(completed)')
    marker='        Case "MVTGaussMixed"'
    assert code.count(marker)==1
    code=code.replace(marker,'        Case "ControlProbe"\n            completed = RunControlProbesTest(report, tracePath)\n        Case "MaterialContexts"\n            completed = RunMaterialContextsTest(report, tracePath)\n        Case "Trace44", "Trace74"\n            completed = RunProductionContinuationTest(CLng(Mid$(caseKey, 6)), report, tracePath)\n'+marker,1)
    code+='\n'+(here/'TestNativeControlProbes.bas.txt').read_text('utf-8')
    files['TestNativeControlProbes.bas.txt']=(here/'TestNativeControlProbes.bas.txt').read_bytes()
    code+='\n'+(here/'TestNativeProductionContinuation.bas.txt').read_text('utf-8')
    files['TestNativeProductionContinuation.bas.txt']=(here/'TestNativeProductionContinuation.bas.txt').read_bytes()
    code+='\n'+(here/'TestNativeMaterialContexts.bas.txt').read_text('utf-8')
    files['TestNativeMaterialContexts.bas.txt']=(here/'TestNativeMaterialContexts.bas.txt').read_bytes()
    contexts=Path('D:/CodexScratch/RC_Section_NDM_TQ/ControlFailureProbe0220261118/Native/ControlProbe.txt').read_bytes()
    assert sha256(contexts).hexdigest().upper()=='93C578BDF09DEFC79521C379671B234224DB7AE44841CC867CB1451498EA50D7'
    contexts=[line for line in contexts.decode('utf-8-sig').splitlines() if line.startswith('MATERIAL_CONTEXT;')]
    assert len(contexts)==5
    files['Data/FailureContexts.csv']=('\r\n'.join(contexts)+'\r\n').encode('cp1251')
    files['tests/modTestMVT2ProductionASDTest.bas']=code.replace('\n','\r\n').encode('utf-8-sig')
    names.append('modTestMVT2ProductionASDTest')
    raw=(mvt/'FinePrefix.csv').read_bytes()
    assert sha256(raw).hexdigest().upper()==mvt_prep['inputs']['FinePrefix.csv']
    files['FinePrefix.csv']=raw
    reference=(repo/'docs/regression/Coupled6D/ProductionGaussAndASD20261118.zip').read_bytes()
    assert sha256(reference).hexdigest().upper()=='322CBF36E5FD026CD3C7A2D852CF3BC0FF1E91A529B5FBBF8CD6186254B96395'
    accepted=[]
    with ZipFile(BytesIO(reference)) as archive:
        for prefix in ('Prefix/Native/','NativeCase/Mechanics/'):
            loads={}
            for line in archive.read(prefix+'MVTGaussMixed.Tensor.csv').decode('cp1251').splitlines():
                row=line.split(';')
                if row[0]=='FRAME': loads[row[1]]=row[2:4]
            for line in archive.read(prefix+'Operators.csv').decode('cp1251').splitlines():
                row=line.split(';')
                if row[0]=='F': accepted.append(';'.join([row[1],*loads[row[1]],*row[2:]]))
    assert len(accepted)==74
    files['AcceptedHistory.csv']=('\r\n'.join(accepted)+'\r\n').encode()
    paths = {
        'Tension': [(-.2*x, -.2*x, x, 0., 0., 0.) for x in (0.,1e-5,5e-5,1e-4,3e-4,.001,.003,.006,.003,0.,-.0003,.002,.01)],
        'Compression': [(-.2*x, -.2*x, x, 0., 0., 0.) for x in (0.,-1e-5,-1e-4,-5e-4,-.001,-.002,-.003,-.005,-.002,0.,.0001,-.001)],
        'Shear': [(0.,0.,0.,x,0.,0.) for x in (0.,1e-5,5e-5,.0001,.0003,.001,.003,.006,.002,0.,-.001,-.003,.001)],
        'Hydrostatic': [(x,x,x,0.,0.,0.) for x in (0.,-1e-5,-.0001,-.0005,-.001,-.0002,0.,1e-5,.0001,0.)],
        'Mixed': [tuple(x*y for y in (-.3,.2,-1.,.17,-.7,.5)) for x in (0.,1e-5,1e-4,.0005,.001,.002,.004,.001,0.,-.0005,-.001,.001)],
    }
    paths['PermutedMixed'] = [(z,x,y,b,c,a) for x,y,z,a,b,c in paths['Mixed']]
    paths['Repeated'] = paths['Mixed'][:8]+[paths['Mixed'][7]]*3+paths['Mixed'][8:]
    def TestRotate(row):
        """Поворачивает инженерные деформации двумя независимыми пространственными углами."""
        x,y,z,xy,xz,yz=row
        tensor=((x,xy/2.,xz/2.),(xy/2.,y,yz/2.),(xz/2.,yz/2.,z))
        c,s,d,t=math.cos(.53),math.sin(.53),math.cos(.37),math.sin(.37)
        rotation=((c*d,-s,c*t),(s*d,c,s*t),(-t,0.,d))
        value=[[sum(rotation[i][k]*tensor[k][m]*rotation[j][m] for k in range(3) for m in range(3)) for j in range(3)] for i in range(3)]
        return value[0][0],value[1][1],value[2][2],2.*value[0][1],2.*value[0][2],2.*value[1][2]
    paths['RotatedCompression']=list(map(TestRotate,paths['Compression']))
    paths['RotatedMixed']=list(map(TestRotate,paths['Mixed']))
    files['Data/History.csv'] = ('\r\n'.join(';'.join([name,str(i),*(format(v,'.17g') for v in row)])
                                         for name, path in paths.items() for i, row in enumerate(path))+'\r\n').encode()
    for name in (Path(__file__).name, 'Run-TestWorkingASD.ps1'):
        files[name] = (here/name).read_bytes()
    root.mkdir()
    for directory in ('Native','Temp','tests'):
        (root/directory).mkdir()
    for name, raw in files.items():
        file = root/name
        file.parent.mkdir(parents=True, exist_ok=True)
        with file.open('xb') as stream: stream.write(raw)
    record = {'base_stage_commit': check_output(['git','rev-parse','HEAD'],cwd=repo).decode().strip(),
              'source_book_sha256':expected,'modules':names,'prepared_tables_not_raw_inputs':True,
              'inputs':{name:sha256(raw).hexdigest().upper() for name,raw in files.items() if name!='docs/regression/Diagnostic.xlsm'}}
    (root/'Preparation.json').write_bytes((json.dumps(record,indent=2,sort_keys=True)+'\n').encode())
    print('WORKING ASD INPUTS READY',len(files),'histories',len(paths),'steps',sum(map(len,paths.values())))


if __name__ == '__main__':
    parser=ArgumentParser(description=__doc__)
    parser.add_argument('--repo',type=Path,default=Path.cwd())
    parser.add_argument('--root',type=Path,required=True)
    args=parser.parse_args()
    TestStage(args.repo.resolve(),args.root.resolve())
