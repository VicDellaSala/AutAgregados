"""Run the real-data pipeline locally; input and output files are never committed."""
import argparse
import json
from pathlib import Path
import sys
import time
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from agregados.procesamiento import process


def main():
    parser = argparse.ArgumentParser(description='Procesar los cuatro archivos sin Streamlit.')
    for name in ['r34', 'diccionario', 'cobro', 'parque', 'periodo', 'salida']:
        parser.add_argument('--'+name, required=True)
    args = parser.parse_args()
    start = time.perf_counter()
    result = process(args.r34, args.diccionario, args.cobro, args.parque, args.periodo, args.salida,
                     progress=lambda n: print(f'{n:,} registros leídos', flush=True))
    report = {'segundos': round(time.perf_counter()-start, 2), 'controles': result.audit.controls(int(result.summary.EQUIPOS.sum())),
              'exclusiones': dict(result.audit.excluded), 'incidencias': dict(result.audit.issues), 'notas': result.audit.notes}
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
