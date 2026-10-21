import io,pathlib,tempfile,unittest,warnings,zipfile
from xlsx_formula_witness import analyze,translate,make_example,NS

def book(cells):
    b=io.BytesIO()
    with zipfile.ZipFile(b,'w') as z:z.writestr('xl/worksheets/sheet1.xml','<worksheet xmlns="'+NS[1:-1]+'"><sheetData><row r="2">'+cells+'</row></sheetData></worksheet>')
    return b.getvalue()
def c(ref,formula):return '<c r="'+ref+'">'+formula+'</c>'

class FormulaTests(unittest.TestCase):
    def test_translation_absolute_strings_function_and_range(self):
        self.assertEqual(translate('SUM(A2:B3)+$A2+A$2+$A$2+"A2"+LOG10(A2)','B2','C3'),'SUM(B3:C4)+$A3+B$2+$A$2+"A2"+LOG10(B3)')
    def test_real_example_roundtrip_and_no_overwrite(self):
        with tempfile.TemporaryDirectory() as t:
            p=pathlib.Path(t)/'demo.xlsx';make_example(p);raw=p.read_bytes();r=analyze(raw)
            self.assertEqual(r['sheets'][0]['formulas'][1]['expanded'],'A3+$A$1');self.assertEqual(r['findings'],[]);self.assertEqual(p.read_bytes(),raw)
            with self.assertRaises(FileExistsError):make_example(p)
    def test_follower_before_master_and_ignored_follower_text(self):
        r=analyze(book(c('B3','<f t="shared" si="0">WRONG</f>')+c('B2','<f t="shared" si="0" ref="B2:B3">A2</f>')))
        f=r['sheets'][0]['formulas'][0];self.assertEqual(f['expanded'],'A3');self.assertTrue(f['source_ignored_by_shared_master'])
    def test_missing_outside_duplicate_and_overlapping_masters(self):
        cases=[c('B3','<f t="shared" si="0"/>'),c('B2','<f t="shared" si="0" ref="B2:B3">A2</f>')+c('B4','<f t="shared" si="0"/>'),c('B2','<f t="shared" si="0" ref="B2:B3">A2</f>')+c('B3','<f t="shared" si="0" ref="B2:B3">A2</f>'),c('B2','<f t="shared" si="0" ref="B2:B4">A2</f>')+c('B3','<f t="shared" si="1" ref="B3:B5">A3</f>')]
        for cells in cases:self.assertTrue(analyze(book(cells))['findings'])
    def test_unsupported_syntax_and_reference_overflow(self):
        for formula in ['Sheet1!A1','Table[Col]','named_range','[Book]Sheet!A1']:
            with self.assertRaises(ValueError):translate(formula,'B2','B3')
        with self.assertRaises(ValueError):translate('A1','B2','A1')
    def test_archive_and_xml_rejections(self):
        with self.assertRaises(ValueError):analyze(book(c('A1','')+c('A1','')))
        raw=book('').replace(b'x',b'x')
        with self.assertRaises(zipfile.BadZipFile):analyze(b'not zip')
        b=io.BytesIO()
        with warnings.catch_warnings():
            warnings.simplefilter('ignore')
            with zipfile.ZipFile(b,'w') as z:
                z.writestr('xl/worksheets/sheet1.xml','x');z.writestr('xl/worksheets/sheet1.xml','y')
        with self.assertRaises(ValueError):analyze(b.getvalue())
        b=io.BytesIO()
        with zipfile.ZipFile(b,'w') as z:z.writestr('xl/worksheets/sheet1.xml','<!DOCTYPE worksheet [<!ENTITY x "y">]><worksheet/>')
        with self.assertRaises(ValueError):analyze(b.getvalue())

if __name__=='__main__':unittest.main()
