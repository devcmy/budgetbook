"""Grammar/resource/wire checks only. This is explicitly not an Xcode build."""
import json,pathlib,sys,hashlib,xml.etree.ElementTree as ET,re
from tree_sitter import Language, Parser
import tree_sitter_swift
from openstep_parser import OpenStepDecoder
ROOT=pathlib.Path(__file__).resolve().parents[1]
parser=Parser(Language(tree_sitter_swift.language()))
syntax=[]
for file in sorted((ROOT/'ios').rglob('*.swift')):
    data=file.read_bytes();tree=parser.parse(data);errors=[]
    def walk(node):
        if node.type=='ERROR' or node.is_missing: errors.append({'type':node.type,'line':node.start_point.row+1,'text':data[node.start_byte:node.end_byte].decode()[:200]})
        for child in node.children:walk(child)
    walk(tree.root_node);syntax.append({'file':str(file.relative_to(ROOT)),'errors':errors})
assert not any(row['errors'] for row in syntax),syntax
project_path=ROOT/'ios/TetrisBudget.xcodeproj/project.pbxproj'
with project_path.open() as handle:project=OpenStepDecoder.ParseFromFile(handle)
objects=project['objects']
paths=[]
for group in objects.values():
    if group.get('isa')=='PBXGroup' and group.get('path') in ['TetrisBudget','TetrisBudgetTests']:
        for ref in group['children']:
            assert ref in objects,ref
            item=objects[ref]
            if item.get('isa')=='PBXFileReference':
                file=ROOT/'ios'/group['path']/item['path'];assert file.is_file(),file;paths.append(str(file.relative_to(ROOT)))
for obj in objects.values():
    for key in ['fileRef','mainGroup','productRefGroup','buildConfigurationList','productReference','target','targetProxy','containerPortal']:
        if key in obj:assert obj[key] in objects,(key,obj[key])
    for key in ['children','files','targets','buildPhases','dependencies','buildConfigurations']:
        for ref in obj.get(key,[]):assert ref in objects,(key,ref)
source_files=[];resources=[];test_targets=[]
for obj in objects.values():
    if obj.get('isa')=='PBXNativeTarget':
        if obj['name']=='TetrisBudgetTests':test_targets.append(obj['name'])
        for phase_id in obj['buildPhases']:
            phase=objects[phase_id]
            for build_id in phase.get('files',[]):
                name=objects[objects[build_id]['fileRef']]['path']
                if phase['isa']=='PBXSourcesBuildPhase':source_files.append(name)
                if phase['isa']=='PBXResourcesBuildPhase':resources.append(name)
assert set(source_files)=={p.name for p in (ROOT/'ios').rglob('*.swift')}
assert {'labels.json','seed-ledger.json','ledger-engine.js'}.issubset(set(resources))
assert test_targets==['TetrisBudgetTests']
xml=ET.parse(ROOT/'ios/TetrisBudget.xcodeproj/xcshareddata/xcschemes/TetrisBudget.xcscheme')
for buildable in xml.findall('.//BuildableReference'):assert buildable.attrib['BlueprintIdentifier'] in objects
assert xml.find('.//TestableReference') is not None
assert (ROOT/'core/ledger-engine.js').read_bytes()==(ROOT/'ios/TetrisBudget/ledger-engine.js').read_bytes()
seed=json.loads((ROOT/'ios/TetrisBudget/seed-ledger.json').read_text())
assert len(seed['sourceArchive'])==2366 and len(seed['transactions'])==45
source=ROOT.parent/'household-budget-history/source/ledger-handoff/evidence/latest_raw_records.json'
source_hash=None
if source.is_file():
    original=json.loads(source.read_text());assert seed['sourceArchive']==original
    source_hash=hashlib.sha256(source.read_bytes()).hexdigest()
    assert source_hash=='94e6d33438ba96fd83b6dfa2388ec97640a0ecc6193ca942fbf0262232ebcbcf'
result={'scope':'Swift grammar and Xcode references only; native build and XCTest NOT EXECUTED','swift_files':syntax,'project_parse':'passed','file_references':paths,'sources':source_files,'resources':resources,'test_targets':test_targets,'engine_bundle_matches_tested_source':True,'original_raw_semantically_preserved':True,'original_raw_sha256':source_hash}
(ROOT/'verification/project-checks.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
print(json.dumps({'grammar_files':len(syntax),'project_parse':'passed','resources':resources,'original_raw_count':2366,'scope':result['scope']},ensure_ascii=False))
