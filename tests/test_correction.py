import unittest
import pandas as pd
from agregados.entradas import Park,load_park
from agregados.facturacion import classify
from agregados.normalizacion import MODES
from test_processing import record,MAPPING,excel

class CorrectionTests(unittest.TestCase):
    def test_posglobal_four_conditions_and_other_exempt_priority(self):
        rows=[record('POSGLOBAL',bank=b,rango=r,tx=t) for b in ['BANCO DE VENEZUELA','MI BANCO'] for r,t in [('30 DIAS SIN TX','0'),('TRANSANDO','1')]]
        rows.append(record('DISGLOBAL',rango='30 DIAS SIN TX',tx='0'))
        result,_=classify(pd.DataFrame(rows),MAPPING,{('POSGLOBAL','BANCO DE VENEZUELA'),('DISGLOBAL','BANCO DE VENEZUELA')},Park())
        self.assertEqual(result.FACTURACION.tolist(),[MODES[0],MODES[0],MODES[2],MODES[1],MODES[2]])

    def test_excluded_canonical_aliases_and_dictionary_alone_not_scope(self):
        mapping=dict(MAPPING,CCR='CREDICARDPOS',SITCA='SITCA',AJENO='AJENO')
        rows=[record(a) for a in ['CCR','CREDICARD POS','SITCA','AJENO','DESCONOCIDO','NOVASERVICES','LATIN PAGOS 2021']]
        result,excluded=classify(pd.DataFrame(rows),mapping,{('SITCA','BANCO DE VENEZUELA')},Park())
        self.assertEqual(result.PERTENENCIA2.tolist(),['SITCA','NOVASERVICES','LATIN PAGOS 2021'])
        self.assertEqual(sum(excluded.values()),4)

    def test_park_exact_id_normalization_fallback_and_no_false_match(self):
        p=load_park(excel([['BANCO','AGREGADO','AFIPOS','AFILIADO','TERMINAL'],[201,'RAPIDPAGO',1231,123,1]],'7A_PARQUE_PARCIAL_ACTIVOS'),MAPPING)
        rows=[record('RAPIDPAGO',afipos='001231.0',NUMPOS='99'),record('RAPIDPAGO',afipos='123'),record('RAPIDPAGO',afipos='999',CODIGO_AFILIADO='123',NUMPOS='1'),record('RAPIDPAGO',afipos='',CODIGO_AFILIADO='00123',NUMPOS='1'),record('SERVICIOS J&G',afipos='1231')]
        result,_=classify(pd.DataFrame(rows),MAPPING,set(),p)
        self.assertEqual(result.FACTURACION.tolist(),[MODES[0],MODES[1],MODES[1],MODES[0],MODES[1]])
        self.assertEqual(result.AFIPOS.iloc[3],'1231')
        p.fallback[('RAPIDPAGO','BANCO DE VENEZUELA','123','1')].add('1232')
        result,_=classify(pd.DataFrame([rows[3]]),MAPPING,set(),p)
        self.assertEqual(result.FACTURACION.iloc[0],MODES[3])

    def test_tx_word_is_not_a_transaction_state(self):
        rows=[record('DISGLOBAL',rango=r,tx='0') for r in ['TRANSANDO','SIN TX EN EL MES ACTUAL','120 DIAS SIN TX','TX DESCONOCIDO']]
        result,_=classify(pd.DataFrame(rows),MAPPING,set(),Park())
        self.assertEqual(result.FACTURACION.tolist(),[MODES[1],MODES[2],MODES[2],MODES[3]])

    def test_sitca_always_proforma_with_aliases_billing_and_invalid_range(self):
        mapping=dict(MAPPING,SITCA='SITCA',SITCA_ALIAS='SITCA')
        rows=[record(a,bank=b,rango=r,tx=t) for a in ['SITCA',' sitca ','SITCA_ALIAS']
              for b in ['BANCO DE VENEZUELA','MI BANCO']
              for r,t in [('TRANSANDO','1'),('30 DIAS SIN TX','0'),('DESCONOCIDO','')]]
        result,excluded=classify(pd.DataFrame(rows),mapping,{('SITCA','BANCO DE VENEZUELA')},Park())
        self.assertEqual(len(result),18)
        self.assertFalse(excluded)
        self.assertEqual(set(result.PERTENENCIA2),{'SITCA'})
        self.assertEqual(set(result.FACTURACION),{'PROFORMA'})
        self.assertTrue(result._ISSUE.str.contains('inconsistentes').any())

    def test_sitca_included_without_billing_or_dictionary_entry(self):
        result,excluded=classify(pd.DataFrame([record('SITCA'),record('SITCA NUEVO')]),MAPPING,set(),Park())
        self.assertEqual(result.PERTENENCIA2.tolist(),['SITCA'])
        self.assertEqual(result.FACTURACION.tolist(),['PROFORMA'])
        self.assertEqual(excluded,{'SITCA NUEVO':1})
