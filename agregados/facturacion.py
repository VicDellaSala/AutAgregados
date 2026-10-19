import pandas as pd
from .normalizacion import norm, compact, bank, identifier, EXCLUDED, EXEMPT, SPECIAL, NO_TX, MODES, COLUMNS, BANK_CODES, REFERENCE_AGENTS, match_id, ALWAYS_PROFORMA


def classify(chunk, mapping, billing, park):
    """Many-to-one lookups only: never merge or expand input rows."""
    data = chunk.fillna('').copy()
    original = data['PERTENENCIA'].map(norm)
    canonical = original.map(mapping)
    data['PERTENENCIA2'] = canonical.fillna(data['PERTENENCIA']).replace('', '(SIN PERTENENCIA)')
    excluded = data['PERTENENCIA2'].map(compact).isin(EXCLUDED) | original.map(compact).isin(EXCLUDED)
    allowed = {compact(mapping.get(a, a)) for a in REFERENCE_AGENTS} | EXEMPT | ALWAYS_PROFORMA | {compact(a) for a, _ in billing}
    excluded |= ~data['PERTENENCIA2'].map(compact).isin(allowed)
    excluded_counts = data.loc[excluded, 'PERTENENCIA2'].value_counts().to_dict()
    data = data.loc[~excluded].copy()
    canonical = canonical.loc[data.index]
    data['_ISSUE'] = ''

    def flag(mask, reason):
        mask = mask.reindex(data.index, fill_value=False)
        data.loc[mask, '_ISSUE'] += reason + '; '

    unknown = ~data['PERTENENCIA2'].map(compact).isin(allowed)
    flag(unknown, 'Pertenencia no reconocida')
    agent = data['PERTENENCIA2'].map(norm)
    banks = data['NOMBRE_BANCO'].map(bank)
    invalid_bank = banks.eq('')
    if 'CODIGO_BANCO' in data:
        codes = data['CODIGO_BANCO'].map(identifier).map(BANK_CODES)
        invalid_bank |= codes.notna() & banks.ne(codes)
    flag(invalid_bank, 'Banco vacío o código incompatible')
    afipos = data['AFIPOS'].map(identifier)
    bad_id = ~afipos.str.fullmatch(r'\d+')
    flag(bad_id, 'AFIPOS vacío o no entero')

    counts = pd.to_numeric(data['CANT_TRANS_ACUM_MES'], errors='coerce')
    rango = data['RANGO'].map(norm)
    no_tx = rango.isin(NO_TX)
    transacting = rango.eq('TRANSANDO')
    flag(~((no_tx & counts.eq(0)) | (transacting & counts.gt(0) & counts.mod(1).eq(0))), 'RANGO y transacciones del mes inconsistentes')
    if 'POS_CON_TRANSACCION' in data:
        pos = pd.to_numeric(data['POS_CON_TRANSACCION'], errors='coerce')
        flag((pos.eq(1) & no_tx) | (pos.eq(0) & transacting), 'POS_CON_TRANSACCION difiere del indicador mensual')

    general = [(a, b) in billing for a, b in zip(agent, banks)]
    data['FACTURACION'] = [MODES[0] if hit else MODES[1] for hit in general]
    special = pd.Series([(a, b) in SPECIAL for a, b in zip(agent, banks)], index=data.index)
    pending = unknown | invalid_bank | bad_id | ~(no_tx | transacting)
    exempt = agent.map(compact).isin(EXEMPT) & no_tx
    exempt &= ~(agent.map(compact).eq('POSGLOBAL') & pd.Series(general, index=data.index))
    data.loc[exempt, 'FACTURACION'] = MODES[2]
    primary = {(a, b, match_id(i)) for a, b, i, _ in park.keys}
    for idx in data.index[special & ~invalid_bank]:
        if not park.usable:
            pending.loc[idx] = True
            data.loc[idx, '_ISSUE'] += park.reason + '; '
            continue
        equipment = match_id(afipos.at[idx])
        if not equipment:
            affiliate = match_id(data.at[idx, 'CODIGO_AFILIADO'])
            terminal = match_id(data.at[idx, 'NUMPOS']) if 'NUMPOS' in data else ''
            matches = park.fallback.get((agent.at[idx], banks.at[idx], affiliate, terminal), set())
            if len(matches) == 1:
                equipment = next(iter(matches))
                data.at[idx, 'AFIPOS'] = equipment
                pending.loc[idx] = not (no_tx.at[idx] or transacting.at[idx])
                data.loc[idx, '_ISSUE'] += 'AFIPOS recuperado por afiliado y terminal del parque; '
        if not equipment:
            continue
        hit = (agent.at[idx], banks.at[idx], equipment) in primary
        data.loc[idx, 'FACTURACION'] = MODES[0] if hit else MODES[1]
    data.loc[pending, 'FACTURACION'] = MODES[3]
    # Explicit business override; retain diagnostic flags without changing PROFORMA.
    data.loc[agent.map(compact).isin(ALWAYS_PROFORMA), 'FACTURACION'] = MODES[1]
    # Group spelling variants without changing the original RNC bank/range fields.
    data['_BANK'] = banks.replace('', '(SIN BANCO)')
    data['_RANGE'] = rango.replace('', '(SIN RANGO)')
    return data, excluded_counts
