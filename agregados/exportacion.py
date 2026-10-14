"""Two-sheet streaming export matching the compact RNC.xlsx reference."""
from pathlib import Path
from xml.sax.saxutils import quoteattr
import xlsxwriter
from .normalizacion import COLUMNS
from .pivot import append_native_pivot, native_layout


class ExcelExport:
    def __init__(self, path, tempdir):
        self.path, self.row, self.closed = path, 1, False
        self.book = xlsxwriter.Workbook(str(path), {'constant_memory': True, 'tmpdir': tempdir,
            'strings_to_formulas': False, 'strings_to_urls': False, 'strings_to_numbers': False})
        self.header = self.book.add_format({'bold': True, 'bg_color': '#D9E1F2'})
        self.summary = self.book.add_worksheet('Hoja1')
        self.summary.set_column('A:A', 32)
        self.summary.set_column('B:B', 24)
        self.summary.activate()
        self.rnc = self.book.add_worksheet('RNC')
        self.rnc.set_column('A:B', 20, self.book.add_format({'num_format': '@'}))
        self.rnc.set_column('C:C', 48)
        self.rnc.set_column('D:M', 28)
        self.rnc.write_row(0, 0, COLUMNS, self.header)
        self.values = [{} for _ in COLUMNS]
        self.records_path = Path(tempdir)/'pivot-records.xml'
        self.records = self.records_path.open('wb')

    def write_rnc(self, data):
        if self.row + len(data) > 1048576:
            raise ValueError('El resultado supera 1.048.575 equipos; Excel no admite más filas en una sola hoja RNC.')
        for values in data.itertuples(index=False, name=None):
            if any(len(str(v)) > 32767 for v in values):
                raise ValueError('Una celda supera el límite de 32.767 caracteres de Excel.')
            if not str(values[0]).strip():
                raise ValueError('Un equipo comercial no tiene AFIPOS y no pudo recuperarse con seguridad del parque.')
            self.rnc.write_row(self.row, 0, values)
            cache_row = []
            for i, (mapping, value) in enumerate(zip(self.values, values)):
                value = str(value)
                if i in (8, 11):
                    index = mapping.setdefault(value, len(mapping))
                    cache_row.append(f'<x v="{index}"/>')
                else:
                    cache_row.append(f'<s v={quoteattr(value)}/>')
            self.records.write(('<r>'+''.join(cache_row)+'</r>').encode())
            self.row += 1

    def finish(self, summary, audit):
        self.rnc.autofilter(0, 0, self.row-1, 12)
        lines = native_layout(summary)
        self.summary.write_row(1, 0, ['FACTURACION', 'PROFORMA' if 'PROFORMA' in self.values[11] else '(Todas)'])
        self.summary.write_row(3, 0, ['Etiquetas de fila', 'Cuenta de AFIPOS'], self.header)
        for row, (agent, count) in enumerate(lines.items(), 4):
            self.summary.write_row(row, 0, [agent, int(count)])
        self.summary.write_row(4+len(lines), 0, ['Total general', int(lines.sum())], self.header)
        self.close()
        append_native_pivot(self.path, summary, self.values, self.records_path, self.row-1)

    def close(self):
        if not self.closed:
            self.records.close()
            self.book.close()
            self.closed = True
