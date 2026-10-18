"""Repair only sheetPr child order; never recalculate or round-trip cell data."""
import argparse
from hashlib import sha256
import json
from pathlib import Path
import posixpath
from zipfile import ZipFile

from lxml import etree as ET
import openpyxl

from validate_excel_structure import validate_package

NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RNS='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
Q=lambda tag:'{'+NS+'}'+tag


def repair(source,output,report,generator_source=None):
    source,output=Path(source),Path(output)
    assert source.resolve()!=output.resolve(),'Repair to a new file before installing it.'
    before=validate_package(source)
    assert before['errors'],'Expected the reported pre-repair structural error.'
    with ZipFile(source) as z:
        infos=z.infolist();parts={i.filename:z.read(i.filename) for i in infos}
    book=ET.fromstring(parts['xl/workbook.xml'])
    sheet=next(s for s in book.find(Q('sheets')) if s.get('name')=='잔액')
    rels=ET.fromstring(parts['xl/_rels/workbook.xml.rels'])
    target=next(r.get('Target') for r in rels if r.get('Id')==sheet.get('{'+RNS+'}id'))
    part=target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/'+target)
    assert part=='xl/worksheets/sheet4.xml'
    assert all(e['part']==part for e in before['errors']),before['errors']
    tree=ET.fromstring(parts[part],ET.XMLParser(resolve_entities=False))
    props=tree.find(Q('sheetPr'))
    previous=[ET.QName(c).localname for c in props]
    assert previous==['tabColor','pageSetUpPr','outlinePr'],previous
    outline,page_setup=props.find(Q('outlinePr')),props.find(Q('pageSetUpPr'))
    props.remove(outline);props.insert(list(props).index(page_setup),outline)
    corrected=[ET.QName(c).localname for c in props]
    fixed=ET.tostring(tree,encoding='UTF-8',xml_declaration=True)
    with ZipFile(output,'w') as z:
        for info in infos:z.writestr(info,fixed if info.filename==part else parts[info.filename])
    after=validate_package(output)
    assert not after['errors'],after['errors']
    with ZipFile(output) as z:
        assert z.testzip() is None
        assert [i.filename for i in infos]==z.namelist()
        changed=[n for n,body in parts.items() if body!=z.read(n)]
        assert changed==[part],changed
        actual=ET.fromstring(z.read(part))
    original=ET.fromstring(parts[part])
    # Child order is the ONLY XML change; cell/formula/cache bytes are exact.
    for a,b in zip(original,actual):
        assert a.tag==b.tag
        if a.tag!=Q('sheetPr'):assert ET.tostring(a)==ET.tostring(b),a.tag
    assert original.attrib==actual.attrib
    oldprops=original.find(Q('sheetPr'));newprops=actual.find(Q('sheetPr'))
    assert oldprops.attrib==newprops.attrib
    assert {c.tag:ET.tostring(c) for c in oldprops}=={c.tag:ET.tostring(c) for c in newprops}
    assert ET.tostring(original.find(Q('sheetData')))==ET.tostring(actual.find(Q('sheetData')))
    values={}
    with_values=openpyxl.load_workbook(output,read_only=True,data_only=True)
    for sn,addresses in {'잔액':['G7','G15','G16','B35','AC501'],'예비비':['L13','F13','I13']}.items():
        values[sn]={addr:with_values[sn][addr].value for addr in addresses}
    assert values['잔액']['G16']==6685466 and values['잔액']['G7']==4535466
    assert values['잔액']['G15']==0 and values['잔액']['AC501']==15
    assert values['예비비']['L13']==239000 and values['예비비']['F13']==0
    generator_check=None
    if generator_source:
        from balance_selected_ledger import prepare
        generated=output.with_name('generator-regression.xlsm')
        prepare(generator_source,generated)
        validation=validate_package(generated)
        assert not validation['errors'],validation['errors']
        with ZipFile(generated) as z:
            properties=ET.fromstring(z.read(part)).find(Q('sheetPr'))
        order=[ET.QName(c).localname for c in properties]
        assert order==['tabColor','outlinePr','pageSetUpPr']
        generator_check={'passed':True,'disposable_generated_file_not_delivered':True,
                         'sheet_properties_order':order,'structure_validation':validation}
    result={'repaired_at':'2026-10-18','fixed_part':part,'sheet':'잔액',
            'cause':'CT_SheetPr requires outlinePr before pageSetUpPr.',
            'before_child_order':previous,'after_child_order':corrected,
            'before_structure_errors':before['errors'],'after_structure_validation':after,
            'changed_zip_parts':changed,'all_cells_formulas_and_cached_values_byte_identical':True,
            'all_other_zip_parts_byte_identical':True,'vba_controls_drawings_styles_byte_identical':True,
            'no_recalculation_or_cell_reserialization':True,'values':values,
            'baseline_sha256':sha256(source.read_bytes()).hexdigest(),
            'workbook_sha256':sha256(output.read_bytes()).hexdigest(),
            'generator_regression':generator_check,
            'validation_scope':'Focused worksheet-properties XSD and structural invariants, not full ECMA validation or native Excel execution.',
            'native_excel_opened':False}
    Path(report).write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps({'repaired':True,'changed_parts':changed,'worksheets':after['worksheet_count'],'errors':len(after['errors']),'hash':result['workbook_sha256']}),flush=True)


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    for name in ['source','output','report']:parser.add_argument('--'+name,required=True)
    parser.add_argument('--generator-source')
    a=parser.parse_args();repair(a.source,a.output,a.report,a.generator_source)
