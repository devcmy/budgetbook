"""Standalone small-package fixtures for validation-tab deletion."""
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile

from lxml import etree as ET

from remove_validation_sheet import ExtendedPackage, remove_validation_sheet
from update_20261015 import NS, PNS, RNS, Q

VNS = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
ANS = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'


class RemoveValidationSheetTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.source = Path(self.directory.name) / 'source.xlsm'
        book = ET.Element(Q('workbook'), nsmap={None: NS, 'r': RNS})
        views = ET.SubElement(book, Q('bookViews'))
        ET.SubElement(views, Q('workbookView'), activeTab='2', firstSheet='1')
        sheets = ET.SubElement(book, Q('sheets'))
        for name, sheetid, rid in [('사용안내', '1', 'rId1'), ('검증', '2', 'rId2'), ('예비비', '9', 'rId9')]:
            ET.SubElement(sheets, Q('sheet'), name=name, sheetId=sheetid,
                          attrib={'{'+RNS+'}id': rid})
        names = ET.SubElement(book, Q('definedNames'))
        for name, local in [('검증', 1), ('예비비', 2)]:
            ET.SubElement(names, Q('definedName'), name='_xlnm.Print_Area', localSheetId=str(local)).text = name+'!$A$1:$C$3'
        rels = ET.Element('{'+PNS+'}Relationships', nsmap={None: PNS})
        for sheetid in ['1', '2', '9']:
            ET.SubElement(rels, '{'+PNS+'}Relationship', Id='rId'+sheetid,
                          Type=RNS+'/worksheet', Target='worksheets/sheet'+sheetid+'.xml')
        ET.SubElement(rels, '{'+PNS+'}Relationship', Id='rIdVBA',
                      Type='http://schemas.microsoft.com/office/2006/relationships/vbaProject', Target='vbaProject.bin')
        app = ET.Element('{'+ANS+'}Properties', nsmap={None: ANS, 'vt': VNS})
        headings = ET.SubElement(ET.SubElement(app, '{'+ANS+'}HeadingPairs'), '{'+VNS+'}vector', size='4', baseType='variant')
        for tag, text in [('lpstr', '워크시트'), ('i4', '3'), ('lpstr', '이름 지정된 범위'), ('i4', '2')]:
            ET.SubElement(ET.SubElement(headings, '{'+VNS+'}variant'), '{'+VNS+'}'+tag).text = text
        titles = ET.SubElement(ET.SubElement(app, '{'+ANS+'}TitlesOfParts'), '{'+VNS+'}vector', size='5', baseType='lpstr')
        for text in ['사용안내', '검증', '예비비', '검증!Print_Area', '예비비!Print_Area']:
            ET.SubElement(titles, '{'+VNS+'}lpstr').text = text
        types = ET.Element('{http://schemas.openxmlformats.org/package/2006/content-types}Types')
        for sheetid in ['1', '2', '9']:
            ET.SubElement(types, '{http://schemas.openxmlformats.org/package/2006/content-types}Override',
                          PartName='/xl/worksheets/sheet'+sheetid+'.xml', ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
        parts = {'xl/workbook.xml': ET.tostring(book), 'xl/_rels/workbook.xml.rels': ET.tostring(rels),
                 'docProps/app.xml': ET.tostring(app), '[Content_Types].xml': ET.tostring(types),
                 'xl/vbaProject.bin': b'fixture-vba-do-not-modify'}
        for sheetid in ['1', '2', '9']:
            sheet = ET.Element(Q('worksheet'), nsmap={None: NS})
            ET.SubElement(sheet, Q('sheetData'))
            parts['xl/worksheets/sheet'+sheetid+'.xml'] = ET.tostring(sheet)
        with ZipFile(self.source, 'w') as z:
            for part, data in parts.items():
                z.writestr(part, data)
        self.p = ExtendedPackage(self.source)

    def cell(self, name='사용안내'):
        data = self.p.sheet(name).find(Q('sheetData'))
        row = ET.SubElement(data, Q('row'), r='1')
        return ET.SubElement(row, Q('c'), r='A1')

    def add_drawing(self, shared=False):
        relationship = ET.Element('{'+PNS+'}Relationships', nsmap={None: PNS})
        ET.SubElement(relationship, '{'+PNS+'}Relationship', Id='rDrawing',
                      Type=RNS+'/drawing', Target='../drawings/drawing1.xml')
        self.p.parts['xl/worksheets/_rels/sheet2.xml.rels'] = ET.tostring(relationship)
        if shared:
            self.p.parts['xl/worksheets/_rels/sheet9.xml.rels'] = ET.tostring(relationship)
        drawing_rels = ET.Element('{'+PNS+'}Relationships', nsmap={None: PNS})
        ET.SubElement(drawing_rels, '{'+PNS+'}Relationship', Id='rImage',
                      Type=RNS+'/image', Target='../media/image1.png')
        self.p.parts['xl/drawings/drawing1.xml'] = b'<drawing/>'
        self.p.parts['xl/drawings/_rels/drawing1.xml.rels'] = ET.tostring(drawing_rels)
        self.p.parts['xl/media/image1.png'] = b'fixture-image'

    def test_remove_preserves_ids_vba_and_adjusts_names_views_titles(self):
        report = remove_validation_sheet(self.p)
        self.assertEqual(report['removed_parts'], ['xl/worksheets/sheet2.xml'])
        self.assertNotIn('검증', self.p.sheets)
        book = self.p.tree('xl/workbook.xml')
        self.assertEqual([(x.get('name'), x.get('sheetId')) for x in book.find(Q('sheets'))], [('사용안내', '1'), ('예비비', '9')])
        self.assertEqual(book.find(Q('definedNames'))[0].get('localSheetId'), '1')
        view = book.find(Q('bookViews'))[0]
        self.assertEqual((view.get('activeTab'), view.get('firstSheet')), ('1', '1'))
        app = self.p.tree('docProps/app.xml')
        ns = {'a': ANS, 'v': VNS}
        self.assertEqual(app.find('a:HeadingPairs/v:vector', ns)[1][0].text, '2')
        self.assertEqual(app.find('a:HeadingPairs/v:vector', ns)[3][0].text, '1')
        self.assertEqual([t.text for t in app.find('a:TitlesOfParts/v:vector', ns)], ['사용안내', '예비비', '예비비!Print_Area'])
        output = Path(self.directory.name) / 'output.xlsm'
        self.p.save(output)
        with ZipFile(output) as z:
            self.assertNotIn('xl/worksheets/sheet2.xml', z.namelist())
            self.assertEqual(z.read('xl/vbaProject.bin'), b'fixture-vba-do-not-modify')

    def test_functional_dependency_fails_before_deletion(self):
        ET.SubElement(self.cell(), Q('f')).text = "SUM('검증'!A1:A3)"
        with self.assertRaisesRegex(ValueError, 'Functional references'):
            remove_validation_sheet(self.p)
        self.assertIn('검증', self.p.sheets)
        self.assertIn('xl/worksheets/sheet2.xml', self.p.parts)

    def test_three_dimensional_dependency_fails(self):
        ET.SubElement(self.cell(), Q('f')).text = 'SUM(사용안내:예비비!A1)'
        with self.assertRaisesRegex(ValueError, 'Functional references'):
            remove_validation_sheet(self.p)

    def test_navigation_links_redirect(self):
        links = ET.SubElement(self.p.sheet('사용안내'), Q('hyperlinks'))
        link = ET.SubElement(links, Q('hyperlink'), ref='A1', location="'검증'!A1")
        formula = ET.SubElement(self.cell(), Q('f'))
        formula.text = 'HYPERLINK("#검증!A1","검증")'
        remove_validation_sheet(self.p)
        self.assertEqual(link.get('location'), "'예비비'!A1")
        self.assertEqual(formula.text, 'HYPERLINK("#\'예비비\'!A1","검증")')

    def test_owned_drawing_and_image_removed(self):
        self.add_drawing()
        report = remove_validation_sheet(self.p)
        for part in ['xl/drawings/drawing1.xml', 'xl/drawings/_rels/drawing1.xml.rels', 'xl/media/image1.png']:
            self.assertIn(part, report['removed_parts'])
            self.assertNotIn(part, self.p.parts)

    def test_shared_drawing_and_image_kept(self):
        self.add_drawing(shared=True)
        report = remove_validation_sheet(self.p)
        self.assertIn('xl/drawings/drawing1.xml', report['retained_shared_ancillary_parts'])
        for part in ['xl/drawings/drawing1.xml', 'xl/drawings/_rels/drawing1.xml.rels', 'xl/media/image1.png']:
            self.assertIn(part, self.p.parts)
        self.assertNotIn('xl/worksheets/_rels/sheet2.xml.rels', self.p.parts)

    def test_shared_strings_remain_unchanged(self):
        sst = ('<sst xmlns="'+NS+'"><si><t>검증 탭과 월간 TRUE/FALSE를 확인하세요.</t></si></sst>').encode('utf-8')
        self.p.parts['xl/sharedStrings.xml'] = sst
        cell = self.cell()
        cell.set('t', 's')
        ET.SubElement(cell, Q('v')).text = '0'
        remove_validation_sheet(self.p)
        self.assertEqual(self.p.parts['xl/sharedStrings.xml'], sst)
        self.assertEqual(cell.get('t'), 'inlineStr')
        self.assertIn('예비비·잔액', ''.join(cell.itertext()))

    def test_global_defined_name_dependency_fails(self):
        names = self.p.tree('xl/workbook.xml').find(Q('definedNames'))
        ET.SubElement(names, Q('definedName'), name='Important').text = '검증!$A$1'
        with self.assertRaisesRegex(ValueError, 'Functional references'):
            remove_validation_sheet(self.p)


if __name__ == '__main__':
    unittest.main()
