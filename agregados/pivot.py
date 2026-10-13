"""Portable native Excel pivot over the freshly computed grouped summary.

OOXML parts use the same layout produced by Excel: AGENTE/BANCO rows,
FACTURACION/RANGO page fields and SUM(EQUIPOS). No original cache is reused.
Only the small RESUMEN XML is parsed; the RNC member is copied as a stream.
"""
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
    agents = summary.groupby('AGENTE').EQUIPOS.sum().sort_values(ascending=False, kind='stable')
    lines = []
    for agent, total in agents.items():
        lines.append((agent, None, int(total)))
        for bank, count in summary.loc[summary.AGENTE.eq(agent)].groupby('BANCO', sort=True).EQUIPOS.sum().items():
            lines.append((agent, bank, int(count)))
    return lines


def append_native_pivot(path, summary, header_row):
    """header_row is one-based. Caller writes the matching visible cells first."""
    if summary.empty:
        return
    values = [sorted(summary.iloc[:, i].unique()) for i in range(4)]
    ids = [{v: i for i, v in enumerate(column)} for column in values]
    lines = native_layout(summary)
    cache = ET.Element(f'{{{S}}}pivotCacheDefinition', {
        f'{{{R}}}id': 'rId1', 'createdVersion': '6', 'refreshedVersion': '6',
        'minRefreshableVersion': '3', 'recordCount': str(len(summary))})
    source = node(cache, 'cacheSource', type='worksheet')
    node(source, 'worksheetSource', ref=f'H8:L{len(summary)+8}', sheet='RESUMEN')
    fields = node(cache, 'cacheFields', count=5)
    for i, name in enumerate(summary.columns):
        field = node(fields, 'cacheField', name=name, numFmtId=0 if i < 4 else 3)
        if i < 4:
            items = node(field, 'sharedItems', count=len(values[i]))
            for value in values[i]:
                node(items, 's', v=value)
        else:
            node(field, 'sharedItems', containsSemiMixedTypes=0, containsString=0,
                 containsNumber=1, containsInteger=1, minValue=int(summary.EQUIPOS.min()), maxValue=int(summary.EQUIPOS.max()))
    records = ET.Element(f'{{{S}}}pivotCacheRecords', count=str(len(summary)))
    for row in summary.itertuples(index=False, name=None):
        record = node(records, 'r')
        for i in range(4):
            node(record, 'x', v=ids[i][row[i]])
        node(record, 'n', v=int(row[4]))
    pivot = ET.Element(f'{{{S}}}pivotTableDefinition', {
        'name': 'ResumenDinamico', 'cacheId': '1', 'dataCaption': 'Valores',
        'updatedVersion': '6', 'minRefreshableVersion': '3', 'createdVersion': '6',
        'useAutoFormatting': '1', 'itemPrintTitles': '1', 'indent': '0',
        'outline': '1', 'outlineData': '1', 'multipleFieldFilters': '0',
        'rowHeaderCaption': 'Etiquetas de fila', 'grandTotalCaption': 'Total general'})
    node(pivot, 'location', ref=f'A{header_row}:B{header_row+len(lines)+1}',
         firstHeaderRow=1, firstDataRow=1, firstDataCol=1, rowPageCount=2, colPageCount=1)
    fields = node(pivot, 'pivotFields', count=5)
    for i in range(5):
        if i == 4:
            node(fields, 'pivotField', dataField=1, numFmtId=3, showAll=0)
            continue
        attrs = {'axis': 'axisRow' if i < 2 else 'axisPage', 'showAll': 0}
        if i >= 2:
            attrs['multipleItemSelectionAllowed'] = 1
        field = node(fields, 'pivotField', **attrs)
        items = node(field, 'items', count=len(values[i])+1)
        for idx in range(len(values[i])):
            node(items, 'item', x=idx)
        node(items, 'item', t='default')
    fields = node(pivot, 'rowFields', count=2)
    node(fields, 'field', x=0); node(fields, 'field', x=1)
    items = node(pivot, 'rowItems', count=len(lines)+1)
    for agent, bank, _ in lines:
        item = node(items, 'i', **({'r': 1} if bank is not None else {}))
        node(item, 'x', v=ids[0][agent] if bank is None else ids[1][bank])
    node(node(items, 'i', t='grand'), 'x')
    node(node(pivot, 'colItems', count=1), 'i')
    pages = node(pivot, 'pageFields', count=2)
    node(pages, 'pageField', fld=2, hier=-1); node(pages, 'pageField', fld=3, hier=-1)
    node(node(pivot, 'dataFields', count=1), 'dataField', name='Cantidad de equipos',
         fld=4, subtotal='sum', baseField=0, baseItem=0, numFmtId=3)
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
        'xl/pivotCache/pivotCacheRecords1.xml': xml(records),
        'xl/pivotTables/_rels/pivotTable1.xml.rels': relationships([('rId1','pivotCacheDefinition','../pivotCache/pivotCacheDefinition1.xml')]),
        'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': relationships([('rId1','pivotCacheRecords','pivotCacheRecords1.xml')]),
    }
    path = Path(path)
    staged = path.with_suffix('.pivot.xlsx')
    try:
        with ZipFile(path) as original:
            sheet = ET.fromstring(original.read('xl/worksheets/sheet2.xml'))
            pivots = node(sheet, 'pivotTableParts', count=1)
            node(pivots, 'pivotTablePart', **{f'{{{R}}}id': 'rIdNativePivot'})
            parts['xl/worksheets/sheet2.xml'] = xml(sheet)
            relname = 'xl/worksheets/_rels/sheet2.xml.rels'
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
                for name, content in parts.items():
                    target.writestr(name, content)
        staged.replace(path)
    finally:
        staged.unlink(missing_ok=True)
