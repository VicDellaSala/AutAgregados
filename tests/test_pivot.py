import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET
import openpyxl
import test_processing as fixtures

N={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


class NativePivotTests(unittest.TestCase):
    def test_native_cache_contains_only_new_weighted_groups(self):
        with tempfile.TemporaryDirectory() as temp:
            output=Path(temp)/'out.xlsx'
            result=fixtures.IntegrationTests().run_process([
                fixtures.record(), fixtures.record(), fixtures.record('PEQUENO'), fixtures.record('CARROPAGO')], output)
            with ZipFile(output) as z:
                self.assertIn(b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"', z.read('[Content_Types].xml'))
                p=ET.fromstring(z.read('xl/pivotTables/pivotTable1.xml'))
                c=ET.fromstring(z.read('xl/pivotCache/pivotCacheDefinition1.xml'))
                records=ET.fromstring(z.read('xl/pivotCache/pivotCacheRecords1.xml'))
                self.assertEqual(c.find('s:cacheSource/s:worksheetSource',N).attrib, {'ref':'H8:L10','sheet':'RESUMEN'})
                self.assertEqual(int(c.get('recordCount')), 2)
                self.assertEqual(sum(int(r.find('s:n',N).get('v')) for r in records), 3)
                self.assertEqual(p.find('s:dataFields/s:dataField',N).get('subtotal'), 'sum')
                self.assertEqual([f.get('x') for f in p.find('s:rowFields',N)], ['0','1'])
                self.assertEqual([f.get('fld') for f in p.find('s:pageFields',N)], ['2','3'])
            # openpyxl validates the native objects and their relationships too.
            w=openpyxl.load_workbook(output)
            self.assertEqual(len(w['RESUMEN']._pivots),1)
            self.assertEqual(w['RNC'].max_row,4)
            self.assertEqual(w['RESUMEN']['B1'].value,'(Todas)')
            self.assertEqual(w['RESUMEN'].auto_filter.ref,'H8:L10')
            w.close()
            fixtures.IntegrationTests().run_process([fixtures.record('PEQUENO')], output)
            with ZipFile(output) as z:
                records=ET.fromstring(z.read('xl/pivotCache/pivotCacheRecords1.xml'))
                self.assertEqual(sum(int(r.find('s:n',N).get('v')) for r in records),1)

    def test_no_invalid_empty_pivot(self):
        with tempfile.TemporaryDirectory() as temp:
            output=Path(temp)/'out.xlsx'
            fixtures.IntegrationTests().run_process([fixtures.record('CARROPAGO')],output)
            with ZipFile(output) as z:
                self.assertFalse(any(n.startswith('xl/pivotTables/') for n in z.namelist()))
