"""Two-sheet streaming Excel; template-equivalent hierarchy, not a native pivot."""
from collections import defaultdict
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name
from .normalizacion import COLUMNS, MODES


def literal(value):
    return '"' + str(value).replace('"', '""') + '"'


def criterion(value):
    return literal('=' + str(value).replace('~', '~~').replace('*', '~*').replace('?', '~?'))


def selected(cell):
    escaped = f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({cell},"~","~~"),"*","~*"),"?","~?")'
    return f'IF({cell}="(Todas)","*","="&{escaped})'


class ExcelExport:
    def __init__(self, path, tempdir):
        self.path, self.row, self.closed = path, 1, False
        self.book = xlsxwriter.Workbook(str(path), {'constant_memory': True, 'tmpdir': tempdir,
            'strings_to_formulas': False, 'strings_to_urls': False, 'strings_to_numbers': False})
        self.header = self.book.add_format({'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#17354D', 'text_wrap': True, 'valign': 'vcenter'})
        self.total = self.book.add_format({'bold': True, 'bg_color': '#DCEBF0', 'num_format': '#,##0'})
        self.number = self.book.add_format({'num_format': '#,##0', 'valign': 'vcenter'})
        self.child = self.book.add_format({'indent': 1})
        self.input = self.book.add_format({'bg_color': '#E9F4FC', 'font_color': '#145AA0', 'border': 1})
        self.note = self.book.add_format({'text_wrap': True, 'valign': 'top', 'font_color': '#536779'})
        self.rnc = self.book.add_worksheet('RNC')
        self.rnc.freeze_panes(1, 2)
        self.rnc.set_row(0, 32)
        self.rnc.set_column('A:A', 18, self.book.add_format({'num_format': '@'}))
        self.rnc.set_column('B:B', 52)
        self.rnc.set_column('C:E', 24)
        self.rnc.set_column('F:F', 28, self.book.add_format({'num_format': '@'}))
        self.rnc.set_column('G:L', 29)
        self.rnc.write_row(0, 0, COLUMNS, self.header)

    def write_rnc(self, data):
        if self.row + len(data) > 1048576:
            raise ValueError('El resultado supera 1.048.575 equipos; Excel no admite más filas en una sola hoja RNC.')
        for values in data.itertuples(index=False, name=None):
            if any(len(str(v)) > 32767 for v in values):
                raise ValueError('Una celda supera el límite de 32.767 caracteres de Excel.')
            self.rnc.write_row(self.row, 0, values)
            self.row += 1

    def finish(self, summary, audit, period):
        self.rnc.autofilter(0, 0, max(1, self.row-1), 11)
        self.rnc.hide_gridlines(2)
        s = self.book.add_worksheet('RESUMEN')
        s.hide_gridlines(2)
        s.freeze_panes(8, 1)
        s.set_column('A:A', 38)
        s.set_column('B:B', 35, self.number)
        s.set_column('C:F', 22, self.number)
        s.set_column('G:G', 3)
        s.set_column('H:K', 29)
        s.set_column('L:L', 16, self.number)
        s.set_column('M:M', 3)
        s.set_column('N:N', 57)
        s.set_column('O:O', 30)
        s.set_column('P:R', 30)
        s.set_column('T:W', 30, None, {'hidden': True})
        s.outline_settings(False, False, True, False)
        cells = {}
        def put(row, col, value, fmt=None, cached=None):
            cells[row, col] = (value, fmt, cached)
        agents = summary.groupby('AGENTE').EQUIPOS.sum().sort_values(ascending=False, kind='stable').index.tolist() if len(summary) else []
        banks = sorted(summary.BANCO.unique())
        ranges = sorted(summary.RANGO.unique())
        options = [MODES, ranges, list(agents), banks]
        for r, (label, values) in enumerate(zip(['FACTURACION', 'RANGO', 'AGENTE', 'BANCO'], options)):
            put(r, 0, label, self.total)
            put(r, 1, '(Todas)', self.input)
            col = xl_col_to_name(19+r)
            for k, value in enumerate(['(Todas)'] + list(values)):
                put(k, 19+r, value)
            s.data_validation(r, 1, r, 1, {'validate': 'list', 'source': f'=$'+col+'$1:$'+col+f'${len(values)+1}', 'error_type': 'stop', 'error_message': 'Seleccione un valor de la lista.'})
        put(4, 0, f'Período: {period}', self.total)
        put(5, 0, 'Filtros B1:B4 actualizan la jerarquía.', self.note)
        put(6, 0, 'Resumen equivalente; no tabla dinámica nativa.', self.note)
        put(0, 7, 'Detalle agrupado para filtros de Excel', self.total)
        put(1, 7, 'Los filtros H:L son independientes de B1:B4.', self.note)
        put(2, 7, 'Subtotal visible del detalle', self.total)
        for c, v in enumerate(['Etiquetas de fila', 'Cuenta de AFIPOS'] + MODES):
            put(7, c, v, self.header)
        for c, v in enumerate(summary.columns):
            put(7, 7+c, v, self.header)
        for r, row in enumerate(summary.itertuples(index=False, name=None), 8):
            for c, v in enumerate(row, 7):
                put(r, c, v, self.number if c == 11 else None)
        end = max(9, 8+len(summary))
        ar, br, mr, rr, nr = [f'${c}$9:${c}${end}' for c in 'HIJKL']
        put(2, 11, f'=SUBTOTAL(109,{nr})', self.total, int(summary.EQUIPOS.sum()))
        s.autofilter(7, 7, max(8, 7+len(summary)), 11)
        row, subtotal_rows = 8, []
        def add_line(r, agent, b, values, subtotal=False):
            fmt = self.total if subtotal else self.number
            put(r, 0, agent if subtotal else b, self.total if subtotal else self.child)
            ac = criterion(agent)
            bc = selected('$B$4') if b is None else criterion(b)
            condition = f'OR($B$3="(Todas)",$B$3={literal(agent)})'
            if b is not None:
                condition = f'AND({condition},OR($B$4="(Todas)",$B$4={literal(b)}))'
            for col, mode in enumerate([None] + MODES, 1):
                mc = selected('$B$1') if mode is None else criterion(mode)
                mode_condition = condition if mode is None else f'AND({condition},OR($B$1="(Todas)",$B$1={literal(mode)}))'
                formula = f'=IF({mode_condition},SUMIFS({nr},{ar},{ac},{br},{bc},{mr},{mc},{rr},{selected("$B$2")}),0)'
                count = int(values.EQUIPOS.sum()) if mode is None else int(values.loc[values.FACTURACION.eq(mode), 'EQUIPOS'].sum())
                put(r, col, formula, fmt, count)
        for agent in agents:
            sub = summary.loc[summary.AGENTE.eq(agent)]
            subtotal_rows.append(row+1)
            add_line(row, agent, None, sub, True)
            row += 1
            for b in sorted(sub.BANCO.unique()):
                add_line(row, agent, b, sub.loc[sub.BANCO.eq(b)])
                s.set_row(row, None, None, {'level': 1})
                row += 1
        put(row, 0, 'Total general', self.header)
        for c, mode in enumerate([None] + MODES, 1):
            letter = xl_col_to_name(c)
            formula = '=SUM(' + ','.join(f'{letter}{r}' for r in subtotal_rows) + ')' if subtotal_rows else '=0'
            value = audit.processed if mode is None else audit.modes.get(mode, 0)
            put(row, c, formula, self.total, value)
        # Controls and diagnostic counts stay on RESUMEN to deliver exactly two sheets.
        put(0, 13, 'VALIDACIONES · todos los registros, sin filtros', self.header)
        r = 2
        for label, value in audit.controls(int(summary.EQUIPOS.sum())).items():
            put(r, 13, label, self.note); put(r, 14, value, self.number); r += 1
        for title, values in [('EXCLUSIONES POR AGENTE', audit.excluded), ('PERTENENCIAS NO RECONOCIDAS', audit.unknown), ('INCIDENCIAS (pueden coincidir en un registro)', audit.issues), ('TOTAL POR BANCO (sin filtros)', summary.groupby('BANCO').EQUIPOS.sum().to_dict())]:
            r += 1; put(r, 13, title, self.header); r += 1
            for label, value in sorted(values.items()):
                put(r, 13, label, self.note); put(r, 14, int(value), self.number)
                if len(label) > 65:
                    s.set_row(r, 15 * ((len(label) + 54) // 55))
                r += 1
        for note in audit.notes:
            r += 1; put(r, 13, note, self.note); s.set_row(r, 60)
        r += 2
        for c, label in enumerate(['AFIPOS', 'PERTENENCIA', 'BANCO', 'FACTURACION', 'INCIDENCIA'], 13):
            put(r, c, label, self.header)
        for sample in audit.samples:
            r += 1
            for c, value in enumerate(sample.values(), 13):
                put(r, c, value, self.note)
            s.set_row(r, 45)
        s.set_row(7, 45)
        for (r, c), (value, fmt, cached) in sorted(cells.items()):
            if cached is not None:
                s.write_formula(r, c, value, fmt, cached)
            else:
                s.write(r, c, value, fmt)
        s.set_landscape()
        s.fit_to_pages(1, 0)
        s.print_area(0, 0, row, 5)
        s.repeat_rows(7)
        self.close()

    def close(self):
        if not self.closed:
            self.book.close()
            self.closed = True
