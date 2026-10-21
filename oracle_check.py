"""Optional independent oracle: install openpyxl separately; not a runtime dependency."""
import pathlib,tempfile,openpyxl
from openpyxl.formula.translate import Translator
from xlsx_formula_witness import translate,make_example
cases=['SUM(A2:B3)+$A2+A$2+$A$2+"A2"+LOG10(A2)','IF(A2>0,A2*2,0)','A2^2+SUM($A$1:A2)','1.2E3+A2','TRUE&A2']
for formula in cases:
    for target in ['B3','C4','D10']:
        assert translate(formula,'B2',target)==Translator('='+formula,origin='B2').translate_formula(target)[1:]
with tempfile.TemporaryDirectory() as tmp:
    path=pathlib.Path(tmp)/'demo.xlsx';make_example(path)
    book=openpyxl.load_workbook(path,data_only=False)
    assert book['Example']['B3'].value=='=A3+$A$1';book.close()
print('openpyxl',openpyxl.__version__,'15 translation comparisons and real workbook read passed')
