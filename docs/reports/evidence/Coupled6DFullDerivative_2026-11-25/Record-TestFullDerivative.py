"""ДЛЯ ТЕСТОВ. Архивирует полную разностную сверку738, сохраняя область доказанного результата и полную цель."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime,timezone
import json

repo=Path.cwd();here=Path(__file__).resolve().parent
root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold738TailFullDifferences_20261125')
proof=json.loads((root/'FullDifferenceProof.json').read_bytes())
assert proof['sample_count']==247 and proof['assembled_vs_pure_residual_bitidentical']
assert proof['source_matrix_bitidentical737'] and proof['accepted_tensor_output_bitidentical_to_verified735']
assert all(s['central']['rank']==17 and not s['unavailable'] and s['baseline_repeated_exactly'] for s in proof['steps'])
out=repo/'docs/reports/evidence/Coupled6DFullDerivative_2026-11-25';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('# Сырые свидетельства сохраняются побитно.\n* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
paths=[]
for name in ['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompiledSourceReuse.json','PointTraceSourceComposition.json',
    'docs/regression/CodeUpdate.json','FullDifferenceProof.json','ParentCodeUpdate.json','ParentCompileClosure.json','Run.ndjson',
    'src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledDeformationContinuation.cls','src/Materials/CCoupledConcreteResponse.cls',
    'tests/modTestCoupledElasticState.bas','tools/build_workbook/Update-IsolatedOptimizerCode.ps1','Compile-TestExternalPoint.ps1',
    'Run-TestPostCrackRawReplay.ps1','Input/PostCrackReplay.bin','Input/PostCrackReplayOptions.bin','Input/PostCrackSourceRawHistory.bin',
    'Stage-TestTailFullDifferences.py']:
    paths.append(root/name)
paths += [p for p in (root/'Native').iterdir() if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
for p in paths:
    assert p.is_file(),str(p);target=out/'738'/p.relative_to(root);target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes(p.read_bytes())
for name in ['Stage-TestTailFullDifferences.py','Check-TestTailFullDifferences.py','Check-TestPersistentHEvidence.py',
    'Read-TestColdRawCapture.py',Path(__file__).name]:(out/name).write_bytes((here/name).read_bytes())
for p in (repo/'tmp/pdfs/ptc20261125').glob('page-*.png'):
    target=out/'PrimaryFormulaVisualCheck'/p.name;target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes(p.read_bytes())
pins={str(p.relative_to(out)).replace('\\','/'):sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
literature=repo/'literature/Coupled6D/NumericalControl'
pdf=literature/'KelleyKeyes1996_ICASE_PseudoTransientContinuation.pdf';assert pdf.read_bytes()[:5]==b'%PDF-'
items={p.name:dict(bytes=p.stat().st_size,sha256=sha256(p.read_bytes()).hexdigest().upper()) for p in [pdf,
    literature/'PETSc-PseudoTransient-BrowserExtract-20261125.json',literature/'PseudoTransientReview20261125.md']}
(literature/'PseudoTransientSources20261125.json').write_text(json.dumps(dict(saved_utc=datetime.now(timezone.utc).isoformat(),
    primary_preprint_url='https://www.cs.odu.edu/~mln/ltrs-pdfs/icase-1996-46.pdf',primary_preprint_pages=19,
    published_paper_doi='10.1137/S0036142996304796',published_version_not_claimed_identical=True,
    visual_formula_check_PDF_pages=[2,3,5],download='Python urllib/OpenSSL verified TLS; Windows Schannel attempts failed',
    files=items),indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',full_derivative738=proof,
    previous_goal_turn_classification='progress:738 all247 complete residual samples available,assembled andpurevalues identical anddeterministic; finite-difference ranks17 atbothh andrelativeerrors1.50e-7/2.56e-9. Verified ranklossisnotassembledJacobianbug here. SavedprimaryKelleyKeyes1996 PDF19pages/visualformulas andPETSc SERsource/manualdiscrepancy.739 bounded nativepseudo-time trial staged/compiling,notpromoted,no physicalcapacity; fullgoalpreserved.')
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for p in (path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json'):p.write_bytes(raw)
print('738 ARCHIVED:',len(pins),'FILES; PRIMARY PDF AND REVIEW HASHED; FULL GOAL ACTIVE')
