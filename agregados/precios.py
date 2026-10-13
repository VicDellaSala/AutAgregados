"""Future pricing configuration: classification never invents monetary amounts."""
import json
from pathlib import Path


def load_prices(path):
    data = json.loads(Path(path).read_text(encoding='utf-8-sig'))
    if not isinstance(data.get('tarifas'), list):
        raise ValueError('La configuración requiere una lista tarifas.')
    for rule in data['tarifas']:
        required = {'agente', 'banco', 'facturacion', 'precio', 'moneda', 'condiciones', 'vigencia_desde', 'vigencia_hasta'}
        if not required <= rule.keys():
            raise ValueError('Tarifa incompleta; faltan criterios o vigencia.')
    return data['tarifas']
