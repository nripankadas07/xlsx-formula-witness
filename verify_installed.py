import hashlib,json,pathlib,subprocess,sys,tempfile
ROOT=pathlib.Path(__file__).absolute().parent
REPO=ROOT.name
PY=pathlib.Path(sys.argv[1]).absolute()
CMD=PY.parent/REPO
def check(args,expected,cwd):
    r=subprocess.run([str(CMD)]+[str(x) for x in args],cwd=cwd,text=True,capture_output=True,timeout=20)
    assert r.returncode==expected,(expected,r.returncode,r.stdout,r.stderr)
    if expected==2:assert r.stderr and not r.stdout
    return json.loads(r.stdout) if r.stdout else None
with tempfile.TemporaryDirectory(prefix='installed-check-') as tmp:
    t=pathlib.Path(tmp)
    if REPO=='line-merge-witness':
        args=[]
        for n in ['base','left','right']:
            p=t/(n+'.txt');p.write_bytes((ROOT/(n+'.txt')).read_bytes());args.append(p)
        before=[p.read_bytes() for p in args]
        out=check(args,0,t);assert out['merged_text']=='FIRST\nmiddle\nLAST\n'
        a=t/'a.txt';b=t/'b.txt';a.write_text('A\n');b.write_text('B\n')
        assert check([args[0],a,b],1,t)['merged_text'] is None
        assert before==[p.read_bytes() for p in args]
    elif REPO=='xlsx-formula-witness':
        p=t/'demo.xlsx';check(['--make-example',p],0,t);raw=p.read_bytes()
        out=check([p],0,t);assert out['sheets'][0]['formulas'][1]['expanded']=='A3+$A$1'
        check(['--make-example',p],2,t);assert p.read_bytes()==raw
        bad=t/'bad.xlsx';bad.write_bytes(b'not a zip');check([bad],2,t)
    else:
        p=t/'example.json';p.write_bytes((ROOT/'example.json').read_bytes());raw=p.read_bytes()
        out=check([p],0,t)
        if REPO=='avro-resolution-witness':
            assert out['compatible_for_all_writer_branches'] and any(x['rule']=='reader_default' for x in out['trace'])
            bad=json.loads(raw);bad['reader']['fields'][0]['type']='boolean'
        elif REPO=='event-time-replay':
            assert [x['sum'] for x in out['emissions']]==[7,11] and out['decisions'][-1]['disposition']=='dropped'
            bad=json.loads(raw);bad['events'].append({'kind':'watermark','time':0})
        else:
            assert [x['lot'] for x in out['allocations']]==['early','later'] and out['totals']['remaining']==2
            bad=json.loads(raw);bad['events'].append({'kind':'issue','id':'i2','day':1,'sku':'widget','quantity':99})
        b=t/'case.json';b.write_text(json.dumps(bad));check([b],2 if REPO=='event-time-replay' else 1,t)
        b.write_text('{"events":[],"events":[]}');check([b],2,t)
        assert p.read_bytes()==raw
    check([t/'missing']*(3 if REPO=='line-merge-witness' else 1),2,t)
    print(json.dumps({'installed_console':True,'outside_checkout':True,'expected_output':True,'meaningful_failure_case':True,'source_unchanged':True}))
