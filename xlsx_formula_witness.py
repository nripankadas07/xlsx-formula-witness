"""Inspect shared-formula groups in bounded transitional OOXML ZIP files, read only."""
import argparse,io,json,pathlib,re,sys,zipfile,xml.etree.ElementTree as ET
NS='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
CELL=re.compile(r'(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]{0,6})\Z')
TOKEN=re.compile(r'"(?:[^"]|"")*"|\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}|[A-Za-z_][A-Za-z0-9_.]*|(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[Ee][+-]?[0-9]+)?|[\s+*/^&%=<>(),:;-]')

def need(ok,message):
    if not ok:raise ValueError(message)

def cell(value):
    m=CELL.fullmatch(value or '');need(m is not None,'invalid A1 coordinate')
    ac,letters,ar,number=m.groups();col=0
    for ch in letters.upper():col=col*26+ord(ch)-64
    row=int(number);need(col<=16384 and row<=1048576,'coordinate outside XLSX bounds')
    return col,row,ac,ar

def letters(col):
    out=''
    while col:col,d=divmod(col-1,26);out=chr(65+d)+out
    return out

def rectangle(value):
    parts=(value or '').split(':');need(len(parts) in (1,2),'invalid shared range')
    a=cell(parts[0]);b=cell(parts[-1]);need(a[0]<=b[0] and a[1]<=b[1],'reversed shared range')
    return a[:2]+b[:2]

def translate(formula,anchor,target):
    need(len(formula)<=8192,'formula exceeds 8192 characters')
    a=cell(anchor);t=cell(target);dx=t[0]-a[0];dy=t[1]-a[1];pos=0;out=[]
    while pos<len(formula):
        m=TOKEN.match(formula,pos);need(m is not None,'unsupported formula syntax: names/tables/sheet/external references are not translated')
        token=m.group();after=formula[m.end():].lstrip();cm=CELL.fullmatch(token)
        if token.startswith('"'):out.append(token)
        elif cm and not after.startswith('('):
            col,row,ac,ar=cell(token);col+=0 if ac else dx;row+=0 if ar else dy
            need(1<=col<=16384 and 1<=row<=1048576,'translated reference outside XLSX bounds')
            out.append(ac+letters(col)+ar+str(row))
        elif re.fullmatch(r'[A-Za-z_][A-Za-z0-9_.]*',token):
            need(after.startswith('(') or token.upper() in ('TRUE','FALSE'),'named ranges are not translated');out.append(token)
        else:out.append(token)
        pos=m.end()
    return ''.join(out)

def analyze(raw):
    need(type(raw) is bytes and len(raw)<=4194304,'ZIP limited to 4 MiB')
    findings=[];reports=[]
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        infos=z.infolist();need(len(infos)<=100,'archive limited to 100 entries')
        names=[i.filename for i in infos];need(len(names)==len(set(names)),'duplicate ZIP entry')
        need(sum(i.file_size for i in infos)<=4194304,'uncompressed archive limited to 4 MiB')
        paths=sorted(n for n in names if re.fullmatch(r'xl/worksheets/[^/]+\.xml',n))
        need(0<len(paths)<=16,'expected 1..16 conventional worksheet parts')
        for path in paths:
            xml=z.read(path).decode('utf-8');need('<!DOCTYPE' not in xml.upper() and '<!ENTITY' not in xml.upper(),'DTD/entity declarations rejected')
            root=ET.fromstring(xml);need(root.tag==NS+'worksheet','only transitional OOXML worksheets supported')
            cells=root.findall(NS+'sheetData/'+NS+'row/'+NS+'c');need(len(cells)<=5000,'worksheet limited to 5000 cells')
            seen=set();groups={};formulas=[]
            for c in cells:
                address=c.get('r');need(re.fullmatch(r'[A-Z]{1,3}[1-9][0-9]{0,6}',address or '') is not None,'canonical cell coordinate required');cell(address);need(address not in seen,'duplicate cell coordinate');seen.add(address)
                fs=c.findall(NS+'f');need(len(fs)<=1,'multiple formula elements in a cell')
                if not fs:continue
                f=fs[0];text=f.text or '';kind=f.get('t','normal')
                rec={'cell':address,'kind':kind,'source':text,'cached_value_present':c.find(NS+'v') is not None}
                formulas.append(rec)
                if kind!='shared':continue
                si=f.get('si');need(si is not None and si.isascii() and si.isdigit() and int(si)<=4294967295,'invalid shared group index')
                si=str(int(si));rec['group']=si
                groups.setdefault(si,[]).append((rec,f.get('ref')))
            ranges=[]
            for si,members in groups.items():
                anchors=[(r,ref) for r,ref in members if ref is not None]
                if len(anchors)!=1:
                    findings.append({'sheet':path,'group':si,'reason':'shared group requires exactly one ref-bearing master'});continue
                master,ref=anchors[0]
                try:rect=rectangle(ref)
                except ValueError as e:findings.append({'sheet':path,'group':si,'reason':str(e)});continue
                ranges.append((si,rect))
                for r,_ in members:
                    x,y,*_=cell(r['cell'])
                    if not(rect[0]<=x<=rect[2] and rect[1]<=y<=rect[3]):
                        findings.append({'sheet':path,'cell':r['cell'],'reason':'shared member outside master range'});continue
                    if not master['source']:
                        findings.append({'sheet':path,'cell':r['cell'],'reason':'empty master formula'});continue
                    try:r['expanded']=translate(master['source'],master['cell'],r['cell']);r['master']=master['cell']
                    except ValueError as e:findings.append({'sheet':path,'cell':r['cell'],'reason':str(e)})
                for r,ref2 in members:
                    if r is not master and r['source']:r['source_ignored_by_shared_master']=True
            for i,(a,r) in enumerate(ranges):
                for b,s in ranges[i+1:]:
                    if max(r[0],s[0])<=min(r[2],s[2]) and max(r[1],s[1])<=min(r[3],s[3]):
                        findings.append({'sheet':path,'groups':[a,b],'reason':'overlapping master ranges'})
            reports.append({'part':path,'formulas':formulas})
    return {'sheets':reports,'findings':findings,'scope':'shared-group structure and lexical A1 translation only; formulas are not evaluated or syntax-validated; cached values may be stale'}

def make_example(path):
    with pathlib.Path(path).open('xb') as target:
        with zipfile.ZipFile(target,'w',zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
            z.writestr('_rels/.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
            z.writestr('xl/workbook.xml','<workbook xmlns="'+NS[1:-1]+'" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Example" sheetId="1" r:id="rId1"/></sheets></workbook>')
            z.writestr('xl/_rels/workbook.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
            z.writestr('xl/worksheets/sheet1.xml','<worksheet xmlns="'+NS[1:-1]+'"><sheetData><row r="2"><c r="A2"><v>7</v></c><c r="B2"><f t="shared" si="0" ref="B2:B3">A2+$A$1</f></c></row><row r="3"><c r="A3"><v>9</v></c><c r="B3"><f t="shared" si="0"/></c></row></sheetData></worksheet>')

def main():
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('input');p.add_argument('--make-example',action='store_true');a=p.parse_args()
    try:
        if a.make_example:make_example(a.input);print(json.dumps({'created':a.input}));return 0
        with pathlib.Path(a.input).open('rb') as f:raw=f.read(4194305)
        out=analyze(raw);print(json.dumps(out,indent=2));return int(bool(out['findings']))
    except (ValueError,OSError,zipfile.BadZipFile,ET.ParseError,RuntimeError) as e:print(str(e),file=sys.stderr);return 2

if __name__=='__main__':sys.exit(main())
