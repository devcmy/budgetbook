"""Build the offline functional preview from the same core/seed as the iOS app."""
from pathlib import Path
import json,hashlib,shutil
ROOT=Path(__file__).resolve().parent
PROJECT=ROOT.parent
seed=json.loads((PROJECT/'ios/TetrisBudget/seed-ledger.json').read_text())
labels=json.loads((PROJECT/'ios/TetrisBudget/labels.json').read_text())
def safe_json(value):
    return json.dumps(value,ensure_ascii=False,separators=(',',':')).replace('<','\\u003c').replace('>','\\u003e').replace('\u2028','\\u2028').replace('\u2029','\\u2029')
html=(ROOT/'index.template.html').read_text().replace('__SEED__',safe_json(seed)).replace('__LABELS__',safe_json(labels))
html=html.replace('<title>','<meta http-equiv="Content-Security-Policy" content="default-src \'none\'; script-src \'unsafe-inline\' \'self\'; style-src \'unsafe-inline\' \'self\'; img-src data:; connect-src \'none\'; base-uri \'none\'; form-action \'none\'"><title>',1)
html=html.replace('../core/ledger-engine.js','ledger-engine.js')
shutil.copyfile(PROJECT/'core/ledger-engine.js',ROOT/'ledger-engine.js')
shutil.copyfile(PROJECT/'core/vendor/xlsx.full.min.js',ROOT/'xlsx.full.min.js')
shutil.copyfile(PROJECT/'core/excel-importer.js',ROOT/'excel-importer.js')
for name in ['ledger-engine.js','excel-importer.js']:
    shutil.copyfile(PROJECT/'core'/name,PROJECT/'ios/TetrisBudget'/name)
shutil.copyfile(PROJECT/'core/vendor/xlsx.full.min.js',PROJECT/'ios/TetrisBudget/xlsx.full.min.js')
shutil.copyfile(PROJECT/'core/vendor/SheetJS-LICENSE.txt',PROJECT/'ios/TetrisBudget/SheetJS-LICENSE.txt')
(ROOT/'index.html').write_text(html)
inline=html.replace('<link rel="stylesheet" href="style.css">','<style>'+(ROOT/'style.css').read_text()+'</style>')
for name in ['ledger-engine.js','xlsx.full.min.js','excel-importer.js','app.js']:
    inline=inline.replace(f'<script src="{name}"></script>','<script>'+(ROOT/name).read_text()+'</script>')
(ROOT/'테트리스_가계부_기능테스트.html').write_text(inline)
(ROOT/'build-manifest.json').write_text(json.dumps(dict(core_sha256=hashlib.sha256((ROOT/'ledger-engine.js').read_bytes()).hexdigest(),native_core_matches=(ROOT/'ledger-engine.js').read_bytes()==(PROJECT/'ios/TetrisBudget/ledger-engine.js').read_bytes(),operational=len(seed['transactions']),preserved_raw=len(seed['sourceArchive']),labels=len(labels),standalone='테트리스_가계부_기능테스트.html'),ensure_ascii=False,indent=2)+'\n')
print(json.dumps(dict(html_bytes=len(inline.encode()),operational=len(seed['transactions']),raw=len(seed['sourceArchive']),shared_core=True),ensure_ascii=False))
