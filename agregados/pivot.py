"""Native compact pivot over individual RNC records, with a fresh portable cache."""
from pathlib import Path
import shutil
import xml.etree.ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'
C = 'http://schemas.openxmlformats.org/package/2006/content-types'


def node(parent, tag, **attrs):
    return ET.SubElement(parent, f'{{{S}}}{tag}', {k: str(v) for k, v in attrs.items()})


def xml(root):
    # OPC readers expect a default namespace on Content_Types and relationships.
    ET.register_namespace('', root.tag[1:].split('}', 1)[0])
    ET.register_namespace('r', R)
    return ET.tostring(root, encoding='utf-8', xml_declaration=True)


def native_layout(summary):
    subset = summary.loc[summary.FACTURACION.eq('PROFORMA')]
    if subset.empty:
        subset = summary
    return subset.groupby('AGENTE').EQUIPOS.sum().sort_values(ascending=False, kind='stable')


def append_native_pivot(path, summary, values, records_path, row_count):
    if not row_count:
        return
    lines = native_layout(summary)
    cache = ET.Element(f'{{{S}}}pivotCacheDefinition', {
        f'{{{R}}}id': 'rId1', 'createdVersion': '6', 'refreshedVersion': '6',
        'minRefreshableVersion': '3', 'recordCount': str(row_count)})
    source = node(cache, 'cacheSource', type='worksheet')
    node(source, 'worksheetSource', ref=f'A1:M{row_count+1}', sheet='RNC')
    fields = node(cache, 'cacheFields', count=13)
    from .normalizacion import COLUMNS
    for i, name in enumerate(COLUMNS):
        field = node(fields, 'cacheField', name=name, databaseField=1)
        items = node(field, 'sharedItems', count=len(values[i]))
        for value in values[i]:
            node(items, 's', v=value)
    pivot = ET.Element(f'{{{S}}}pivotTableDefinition', {
        'name': 'TablaDinámica1', 'cacheId': '1', 'dataCaption': 'Valores',
        'updatedVersion': '6', 'minRefreshableVersion': '3', 'createdVersion': '6',
        'useAutoFormatting': '1', 'itemPrintTitles': '1', 'indent': '0',
        'outline': '1', 'outlineData': '1', 'rowHeaderCaption': 'Etiquetas de fila',
        'grandTotalCaption': 'Total general'})
    node(pivot, 'location', ref=f'A4:B{len(lines)+5}', firstHeaderRow=1,
         firstDataRow=1, firstDataCol=1, rowPageCount=1, colPageCount=1)
    fields = node(pivot, 'pivotFields', count=13)
    for i in range(13):
        attrs = {'showAll': 0}
        if i == 0:
            attrs['dataField'] = 1
        if i in (8, 11):
            attrs['axis'] = 'axisRow' if i == 8 else 'axisPage'
        field = node(fields, 'pivotField', **attrs)
        if i in (8, 11):
            items = node(field, 'items', count=len(values[i])+1)
            for idx in range(len(values[i])):
                node(items, 'item', x=idx)
            node(items, 'item', t='default')
    node(node(pivot, 'rowFields', count=1), 'field', x=8)
    items = node(pivot, 'rowItems', count=len(lines)+1)
    for agent in lines.index:
        node(node(items, 'i'), 'x', v=values[8][agent])
    node(node(items, 'i', t='grand'), 'x')
    node(node(pivot, 'colItems', count=1), 'i')
    page = {'fld': 11, 'hier': -1}
    if 'PROFORMA' in values[11]:
        page['item'] = values[11]['PROFORMA']
    node(node(pivot, 'pageFields', count=1), 'pageField', **page)
    node(node(pivot, 'dataFields', count=1), 'dataField', name='Cuenta de AFIPOS',
         fld=0, subtotal='count', baseField=0, baseItem=0)
    node(pivot, 'pivotTableStyleInfo', name='PivotStyleLight16', showRowHeaders=1,
         showColHeaders=1, showRowStripes=0, showColStripes=0, showLastColumn=1)

    def relationships(entries):
        root = ET.Element(f'{{{P}}}Relationships')
        for id_, kind, target in entries:
            ET.SubElement(root, f'{{{P}}}Relationship', Id=id_, Type=R+'/'+kind, Target=target)
        return xml(root)

    parts = {
        'xl/pivotTables/pivotTable1.xml': xml(pivot),
        'xl/pivotCache/pivotCacheDefinition1.xml': xml(cache),
        'xl/pivotTables/_rels/pivotTable1.xml.rels': relationships([('rId1','pivotCacheDefinition','../pivotCache/pivotCacheDefinition1.xml')]),
        'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': relationships([('rId1','pivotCacheRecords','pivotCacheRecords1.xml')]),
    }
    path = Path(path)
    staged = path.with_suffix('.pivot.xlsx')
    try:
        with ZipFile(path) as original:
            sheet = ET.fromstring(original.read('xl/worksheets/sheet1.xml'))
            pivots = node(sheet, 'pivotTableParts', count=1)
            node(pivots, 'pivotTablePart', **{f'{{{R}}}id': 'rIdNativePivot'})
            parts['xl/worksheets/sheet1.xml'] = xml(sheet)
            relname = 'xl/worksheets/_rels/sheet1.xml.rels'
            rels = ET.fromstring(original.read(relname)) if relname in original.namelist() else ET.Element(f'{{{P}}}Relationships')
            ET.SubElement(rels, f'{{{P}}}Relationship', Id='rIdNativePivot', Type=R+'/pivotTable', Target='../pivotTables/pivotTable1.xml')
            parts[relname] = xml(rels)
            book = ET.fromstring(original.read('xl/workbook.xml'))
            caches = ET.Element(f'{{{S}}}pivotCaches')
            node(caches, 'pivotCache', cacheId=1, **{f'{{{R}}}id': 'rIdNativeCache'})
            # pivotCaches follows calcPr and precedes extLst in CT_Workbook.
            ext = book.find(f'{{{S}}}extLst')
            book.insert(list(book).index(ext) if ext is not None else len(book), caches)
            parts['xl/workbook.xml'] = xml(book)
            rels = ET.fromstring(original.read('xl/_rels/workbook.xml.rels'))
            ET.SubElement(rels, f'{{{P}}}Relationship', Id='rIdNativeCache', Type=R+'/pivotCacheDefinition', Target='pivotCache/pivotCacheDefinition1.xml')
            parts['xl/_rels/workbook.xml.rels'] = xml(rels)
            types = ET.fromstring(original.read('[Content_Types].xml'))
            for filename, kind in [('xl/pivotTables/pivotTable1.xml','pivotTable'), ('xl/pivotCache/pivotCacheDefinition1.xml','pivotCacheDefinition'), ('xl/pivotCache/pivotCacheRecords1.xml','pivotCacheRecords')]:
                ET.SubElement(types, f'{{{C}}}Override', PartName='/'+filename, ContentType=f'application/vnd.openxmlformats-officedocument.spreadsheetml.{kind}+xml')
            parts['[Content_Types].xml'] = xml(types)
            with ZipFile(staged, 'w', ZIP_DEFLATED, allowZip64=True) as target:
                for member in original.infolist():
                    if member.filename not in parts:
                        with original.open(member) as src, target.open(member, 'w') as dst:
                            shutil.copyfileobj(src, dst, length=1024*1024)
                with target.open('xl/pivotCache/pivotCacheRecords1.xml', 'w') as dst:
                    dst.write(f'<pivotCacheRecords xmlns="{S}" count="{row_count}">'.encode())
                    with open(records_path, 'rb') as src:
                        shutil.copyfileobj(src, dst, length=1024*1024)
                    dst.write(b'</pivotCacheRecords>')
                for name, content in parts.items():
                    target.writestr(name, content)
        staged.replace(path)
    finally:
        staged.unlink(missing_ok=True)
