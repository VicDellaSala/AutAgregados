from collections import Counter
from dataclasses import dataclass
from pathlib import Path
import re
import tempfile
import pandas as pd
from .normalizacion import COLUMNS, norm, identifier, bank
from .entradas import load_dictionary, load_billing, load_park, park_period_issue
from .facturacion import classify
from .validacion import Audit, DuplicateTracker
from .exportacion import ExcelExport

REQUIRED = set(COLUMNS) - {'PERTENENCIA2', 'FACTURACION'} | {'CANT_TRANS_ACUM_MES', 'ANO_PROCESO', 'MES_PROCESO'}
OPTIONAL = {'NUMPOS', 'CODIGO_BANCO', 'POS_CON_TRANSACCION'}
SUMMARY_COLUMNS = ['AGENTE', 'BANCO', 'FACTURACION', 'RANGO', 'EQUIPOS']


@dataclass
class Result:
    output: Path
    preview: pd.DataFrame
    summary: pd.DataFrame
    audit: Audit
    period: str


def process(r34, dictionary, billing, park, period, output, *, encoding='utf-8-sig', chunksize=25000, progress=None):
    if not re.fullmatch(r'20\d{2}-(0[1-9]|1[0-2])', period):
        raise ValueError('Período inválido; use AAAA-MM.')
    year, month = map(int, period.split('-'))
    mapping, invalid = load_dictionary(dictionary)
    pairs, billing_duplicates = load_billing(billing, mapping)
    park_name = getattr(park, 'name', str(park))
    park_issue = park_period_issue(park_name, period) if park is not None else ''
    park_data = load_park(park, mapping, not bool(park_issue), park_issue)
    audit = Audit()
    if invalid:
        audit.notes.append(f'{len(invalid)} equivalencias vacías o #N/A: se conservan como pendientes, no son agentes válidos.')
    if billing_duplicates or park_data.duplicates:
        audit.notes.append(f'Claves repetidas de referencias colapsadas sin multiplicar filas: cobro={billing_duplicates}, parque={park_data.duplicates}.')
    if not park_data.usable:
        audit.notes.append(park_data.reason)
    audit.notes.append('Sin acumulación histórica. El parque sin fecha interna se asume del período seleccionado, salvo incompatibilidad del nombre.')
    audit.notes.append('POS_CON_TRANSACCION no determina exoneración: se contrasta CANT_TRANS_ACUM_MES con RANGO. Sus discrepancias se reportan.')
    if hasattr(r34, 'seek'):
        r34.seek(0)
        sample = r34.read(16384)
        r34.seek(0)
    else:
        with open(r34, 'rb') as f:
            sample = f.read(16384)
    if isinstance(sample, bytes):
        sample = sample.decode(encoding)
    first = sample.splitlines()[0] if sample else ''
    separator = max([';', ',', '\t'], key=first.count)
    names = [norm(v.strip('"')) for v in first.lstrip('\ufeff').split(separator)]
    if len(names) != len(set(names)):
        raise ValueError('R34 contiene encabezados duplicados.')
    missing = REQUIRED - set(names)
    if missing:
        raise ValueError('Faltan columnas en R34: ' + ', '.join(sorted(missing)))
    usecols = [i for i, value in enumerate(names) if value in REQUIRED | OPTIONAL]
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    counts, preview = Counter(), []
    preview_size = 0
    with tempfile.TemporaryDirectory(prefix='agregados-') as temp:
        tracker = DuplicateTracker(Path(temp) / 'duplicates.sqlite')
        export = ExcelExport(Path(temp) / 'output.xlsx', temp)
        try:
            reader = pd.read_csv(r34, sep=separator, encoding=encoding, dtype=str, keep_default_na=False,
                                 usecols=usecols, chunksize=chunksize, on_bad_lines='error')
            with reader:
                for chunk in reader:
                    chunk.columns = [norm(c) for c in chunk.columns]
                    if not (pd.to_numeric(chunk.ANO_PROCESO, errors='coerce').eq(year) &
                            pd.to_numeric(chunk.MES_PROCESO, errors='coerce').eq(month)).all():
                        raise ValueError(f'El R34 contiene registros de otro período o sin fecha. Se seleccionó {period}.')
                    audit.read += len(chunk)
                    audit.duplicate_keys += tracker.add(zip(chunk.NOMBRE_BANCO.map(bank), chunk.AFIPOS.map(identifier)))
                    data, excluded = classify(chunk, mapping, pairs, park_data)
                    audit.excluded.update(excluded)
                    audit.processed += len(data)
                    audit.modes.update(data.FACTURACION)
                    unknown = data._ISSUE.str.contains('Pertenencia no reconocida', regex=False)
                    audit.unknown.update(data.loc[unknown, 'PERTENENCIA'].replace('', '(VACÍA)'))
                    for reasons, count in data.loc[data._ISSUE.ne(''), '_ISSUE'].value_counts().items():
                        for reason in reasons.rstrip('; ').split('; '):
                            audit.issues[reason] += int(count)
                    if len(audit.samples) < 300:
                        audit.samples.extend(data.loc[data._ISSUE.ne(''), ['AFIPOS', 'PERTENENCIA', 'NOMBRE_BANCO', 'FACTURACION', '_ISSUE']].head(300-len(audit.samples)).to_dict('records'))
                    grouped = data.groupby(['PERTENENCIA2', '_BANK', 'FACTURACION', '_RANGE'], dropna=False).size()
                    counts.update({key: int(value) for key, value in grouped.items()})
                    export.write_rnc(data[COLUMNS])
                    if preview_size < 1000:
                        part = data[COLUMNS].head(1000-preview_size)
                        preview.append(part)
                        preview_size += len(part)
                    if progress:
                        progress(audit.read)
            audit.validate(sum(counts.values()))
            summary = pd.DataFrame([(*key, count) for key, count in sorted(counts.items())], columns=SUMMARY_COLUMNS)
            export.finish(summary, audit, period)
            # Copy only a completed, reconciled workbook to the caller's path.
            import shutil
            shutil.copyfile(export.path, output)
        finally:
            tracker.close()
            export.close()
    return Result(output, pd.concat(preview, ignore_index=True) if preview else pd.DataFrame(columns=COLUMNS), summary, audit, period)
