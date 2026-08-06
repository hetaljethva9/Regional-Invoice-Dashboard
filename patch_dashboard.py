import re, pathlib

src = pathlib.Path("dashboard_source.html")
out = pathlib.Path("index.html")
s = src.read_text(encoding="utf-8", errors="surrogateescape")

# 1) data loader that fetches ./data/invoices.json
m = re.search(r'async function (\w+)\(\)\{try\{const e=await fetch\("\./data/invoices\.json"\);if\(!e\.ok\)throw new Error\(`Failed to load invoice data \(\$\{e\.status\}\)`\);const t=await e\.json\(\);', s)
assert m, "loader pattern not found"
loader, old_loader = m.group(1), m.group(0)

# 2) built-in Excel parser + module loader names
m2 = re.search(r'async function (\w+)\(e,t\)\{const a=await (\w+)\(\(\)=>Promise\.resolve\(\)\.then\(\(\)=>(\w+)\),void 0,import\.meta\.url\),n=await e\.arrayBuffer\(\)', s)
assert m2, "Excel parser not found - this build has no Excel import feature"
parser, shim, xmod = m2.groups()

# 3) mapping merge + builder names
m3 = re.search(r'function (\w+)\(e\)\{const t=new Map\((\w+)\(\)\);for\(const\[a,n\]of e\)t\.set\(a,n\)', s)
assert m3, "mapping merge not found"
merge = m3.group(1)
m4 = re.search(re.escape(parser) + r'\(e,t\)\{.{0,700}?throw new \w+\((\w+)\(o\)\)', s)
assert m4, "mapping builder not found"
builder = m4.group(1)

new_loader = (
 'async function ' + loader + '(){'
 'try{const mf=await fetch("./Mapping.xlsx?v="+Date.now());'
 'if(mf.ok){const ma=await ' + shim + '(()=>Promise.resolve().then(()=>' + xmod + '),void 0,import.meta.url),'
 'mb=await mf.arrayBuffer(),mi=ma.read(mb,{cellDates:!0}),ms=mi.Sheets[mi.SheetNames[0]],'
 'mo=ma.utils.sheet_to_json(ms,{defval:null});mo.length&&' + merge + '(' + builder + '(mo))}}catch{}'
 'try{let xn=null;'
 'try{const lr=await fetch("https://api.github.com/repos/hetaljethva9/Regional-Invoice-Dashboard/contents/?ref=main&t="+Date.now());'
 'if(lr.ok){const ls=await lr.json();'
 'const cand=(Array.isArray(ls)?ls:[]).filter(f=>f.type==="file"&&/^ops_invoicing_merged_.*\\.xlsx$/i.test(f.name)).map(f=>f.name).sort();'
 'xn=cand.length?cand[cand.length-1]:null}}catch{}'
 'const fn=xn||"Test data.xlsx";'
 'const e=await fetch("./"+encodeURIComponent(fn)+"?v="+Date.now());'
 'if(!e.ok)throw new Error(`Failed to load invoice data (${e.status})`);'
 'const t=await ' + parser + '(e);t.__src=fn+" (live)";'
)
assert s.count(old_loader) == 1
s = s.replace(old_loader, new_loader)

# 4) header label
m5 = re.search(r'const (\w+)=await ' + re.escape(loader) + r'\(\);(\w+)\|\|(\w+)\((\w+)\(\1\),"Bundled export"\)', s)
assert m5, "label pattern not found"
v, guard, setter, wrap = m5.groups()
s = s.replace(m5.group(0), f'const {v}=await {loader}();{guard}||{setter}({wrap}({v}),{v}.__src||"Bundled export")')

out.write_text(s, encoding="utf-8", errors="surrogateescape")
print("patched OK")
