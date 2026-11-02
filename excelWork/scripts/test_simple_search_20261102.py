"""No native Excel execution: CFB, MS-OVBA, predicates and formula construction."""
from datetime import datetime
from io import BytesIO
from pathlib import Path
import random
import unittest

import olefile
from oletools.olevba import decompress_stream

from simple_search_20261102 import matches,result_formulas,MODULE
from vba_search_project_20261102 import compress,compound


def vba_string_expression(expression,variables):
    """Evaluate only the quoted-string concatenation used for .Formula writes."""
    pieces=[];current='';quoted=False;i=0
    while i<len(expression):
        c=expression[i]
        if c=='"':
            current+=c
            if quoted and i+1<len(expression) and expression[i+1]=='"':
                current+='"';i+=2;continue
            quoted=not quoted
        elif c=='&' and not quoted:pieces.append(current.strip());current=''
        else:current+=c
        i+=1
    assert not quoted
    pieces.append(current.strip());output=''
    for part in pieces:
        if part in variables:output+=variables[part]
        else:
            assert part.startswith('"') and part.endswith('"'),part
            output+=part[1:-1].replace('""','"')
    return output


class SimpleSearchTests(unittest.TestCase):
    def test_compressed_source_roundtrip(self):
        for data in [b'hello',b'a'*10000,MODULE.read_bytes(),random.Random(13).randbytes(4096)]:
            self.assertEqual(decompress_stream(bytearray(compress(data))),data)

    def test_compound_small_and_large_streams(self):
        streams={('PROJECT',):b'a'*80,('VBA','dir'):b'b'*512,('VBA','Module'):b'c'*6000,('VBA','empty'):b''}
        o=olefile.OleFileIO(BytesIO(compound(streams)))
        self.assertEqual({tuple(p):o.openstream(p).read() for p in o.listdir()},streams)

    def test_search_all_months_default(self):
        row=[datetime(2026,10,23),None,'Card',None,None,'숙소비',169000,None,'화담숲 글램핑 숙소 예약']
        self.assertTrue(matches(row,'글램핑',0))
        self.assertFalse(matches(row,'글램핑',11))
        self.assertTrue(matches(row,'글램핑',10))
        self.assertTrue(matches(row,'169000'))

    def test_literal_and_case_insensitive_search(self):
        row=[datetime(2026,11,2),None,'Card ABC',None,None,'A*B',4000,None,'memo']
        self.assertTrue(matches(row,'abc'))
        self.assertTrue(matches(row,'*'))
        self.assertFalse(matches(row,'A?B'))

    def test_vba_generates_the_expected_live_formulas(self):
        expressions=[line.split('.Formula = ',1)[1] for line in MODULE.read_text().splitlines() if '.Formula = ' in line]
        self.assertEqual(len(expressions),7)
        for row in [8,85,4007]:
            p=f'IFERROR(MATCH($H{row},RAW!$AP$2:$AP$4001,0),0)'
            variables={'position':p,'sourceRow':f'({p}+1)',
                       'view.Range("J3").Value2':'원본 수정','view.Range("J4").Value2':'메모 수정'}
            self.assertEqual([vba_string_expression(e,variables) for e in expressions],result_formulas(row))

    def test_no_financial_posting_or_external_execution(self):
        code=MODULE.read_text()
        self.assertNotIn('raw.Cells',code)
        self.assertNotIn('raw.Range("A2:AQ4001").Value2 =',code)
        self.assertNotIn('Workbook_Open',code)
        for banned in ['Shell(', 'CreateObject(', 'Kill ', 'SaveAs', 'VBProject', 'Declare ', 'Environ(']:
            self.assertNotIn(banned,code)

    def test_search_restores_application_state(self):
        code=MODULE.read_text()
        for prop,value in [('Calculation','savedCalc'),('EnableEvents','savedEvents'),('ScreenUpdating','savedScreen'),('StatusBar','oldStatus')]:
            self.assertIn(f'Application.{prop} = {value}',code)


if __name__=='__main__':unittest.main()
