"""Run the independent/formula tests, then cache the verified baseline for preview."""
from pathlib import Path
import sys, json, unittest, hashlib
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'tests'))
from test_workbook import WorkbookTests, PATH
from process_model import DATA, reference_result
suite=unittest.defaultTestLoader.loadTestsFromTestCase(WorkbookTests)
result=unittest.TextTestRunner(verbosity=2).run(suite)
if not result.wasSuccessful():raise SystemExit(1)
e=WorkbookTests.engine.calculate()
assert not e.errors()
formulas,charts=e.embed()
import openpyxl
assert openpyxl.load_workbook(PATH).calculation.calcMode == 'auto'
manifest={
 'status':'passed',
 'test_cases':result.testsRun,
 'formula_count':formulas,
 'chart_count':charts,
 'formula_engine':'Python formulas (not a GUI Excel/LibreOffice test)',
 'checks':'Independent arithmetic model, whole-cycle exclusions, formula propagation, zero/invalid inputs, exact monoproduct capacity, sanitary event counting, source/target separation, audit completion and resource guards',
 'demo_before':reference_result(),
 'demo_target':reference_result(True),
 'source_commit':DATA['source_commit'],
 'sources':{p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in sorted((ROOT/'sources').iterdir()) if p.is_file()},
 'workbook_sha256':hashlib.sha256(PATH.read_bytes()).hexdigest(),
}
(ROOT/'data/verification.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2))
print(f'RELEASE VERIFIED: {result.testsRun} tests; {formulas} formula caches; {charts} chart caches.')
