import re
import unicodedata


def text(value):
    return '' if value is None else str(value).strip()


def norm(value):
    value = unicodedata.normalize('NFKD', text(value))
    return ' '.join(''.join(c for c in value if not unicodedata.combining(c)).upper().split())


def compact(value):
    return norm(value).replace(' ', '')


def identifier(value):
    """Never pass identifiers through floating point; retain leading zeroes."""
    if isinstance(value, float):
        if not value.is_integer() or abs(value) >= 10**15:
            raise ValueError('Identificador numérico de Excel impreciso; debe guardarse como texto.')
        return str(int(value))
    if isinstance(value, int) and abs(value) >= 10**15:
        raise ValueError('Identificador largo de Excel debe guardarse como texto.')
    return re.sub(r'\.0+$', '', text(value))


BANK_CODES = {
    '201': 'BANCO DE VENEZUELA', '301': 'BANCO DEL CARIBE',
    '501': '501-BICENTENARIO', '805': 'BANPLUS', '810': 'BANCO EXTERIOR',
    '830': 'BANCO DEL TESORO', '840': 'AGRICOLA', '850': 'BANCRECER',
    '860': 'ACTIVO', '880': 'MI BANCO', '885': '100% BANCO',
    '895': 'BANCAMIGA', '896': 'BANFANB', '915': 'B.O.D.',
    '945': 'BANCO PLAZA', '965': 'BANCO CARONI',
    '980': 'BANCO FONDO COMUN', '990': 'DEL SUR',
}


def bank(value):
    value = norm(value)
    value = BANK_CODES.get(identifier(value), value)
    return {'BANFAB': 'BANFANB', 'BANFAN': 'BANFANB'}.get(value, value)


EXCLUDED = {'CREDICARDPOS', 'SITCA', 'CARROPAGO', 'POSITIVOPLUS', 'ACIERTO TECNOLOGICO'.replace(' ', ''), 'AGILPAGOS', 'AGILPAGO'}
EXEMPT = {'DISGLOBAL', 'UBIIPAGOS', 'POSGLOBAL', 'LATINPAGOS2021'}
SPECIAL = {('RAPIDPAGO', 'BANCO DE VENEZUELA'), ('RAPIDPAGO', 'BANFANB'), ('SERVICIOS J&G', 'BANCO DE VENEZUELA')}
NO_TX = {'SIN TX EN EL MES ACTUAL', '30 DIAS SIN TX', '60 DIAS SIN TX', '90 DIAS SIN TX', '120 DIAS SIN TX'}
MODES = ['COBRO CENTRALIZADO', 'PROFORMA', 'EXONERADO', 'PENDIENTE DE VALIDACION']
COLUMNS = ['AFIPOS', 'CODIGO_AFILIADO', 'TERMINAL', 'NOMBRE_BANCO', 'MARCA', 'MODELO', 'SERIAL', 'PERTENENCIA', 'PERTENENCIA2', 'ESTADO_POSV2', 'MANEJADOR', 'FACTURACION', 'RANGO']

# Commercial scope observed in RNC.xlsx / Hoja1, supplemented at runtime by billing.
REFERENCE_AGENTS = {'DISGLOBAL', 'VEPAGOS', 'FASTPAYMENT', 'RAPIDPAGO', 'VA TECHNOLOGY', 'POSGLOBAL', 'INVERSIONES MASTER 2021', 'FG METROPOLIS', 'BEST PAY', 'INTELIGENSA', 'SERVICIOS J&G', 'CORPOCREDIT', 'PAGUETODO', 'TOKEN PAGO', 'CITYPOS', 'KAFREGA', 'POSCOMERCIAL', 'SUNMI VENEZUELA', 'THUNDERNET', 'PAGONET', 'CONSULTEL', 'INSTAPAGO', 'CLEA', 'PAYTECH', 'SERVIPUNTO', 'UNIDIGITAL', 'DIGIPAGO', 'UBIIPAGOS', 'NOVASERVICES'}

def match_id(value):
    value = identifier(value)
    return str(int(value)) if value.isdigit() else ''
