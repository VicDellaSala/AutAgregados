"""Optional Streamlit smoke test; runs when requirements.txt is installed."""
import importlib.util
import tempfile
import io
from unittest.mock import patch
import pandas as pd
import unittest
from pathlib import Path
import test_processing as fixtures


@unittest.skipUnless(importlib.util.find_spec('streamlit'), 'Instale requirements.txt para probar Streamlit')
class StreamlitTests(unittest.TestCase):
    def test_uploads_results_filters_and_stale_result_reset(self):
        from streamlit.testing.v1 import AppTest
        at = AppTest.from_file(str(Path(__file__).resolve().parents[1] / 'app.py'), default_timeout=20).run()
        self.assertFalse(at.exception)
        self.assertEqual(len(at.get('file_uploader')), 4)
        self.assertEqual(len(at.date_input), 0)
        self.assertTrue(at.button[0].disabled)
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        result = fixtures.IntegrationTests().run_process([fixtures.record(), fixtures.record('PEQUENO')], Path(temp.name)/'out.xlsx')
        at.session_state['result'] = {'data': result, 'temp': temp}
        at.run()
        self.assertFalse(at.exception)
        self.assertEqual(len(at.get('download_button')), 1)
        at.multiselect[0].select('PROFORMA').run()
        self.assertFalse(at.exception)
        self.assertEqual(at.metric[-1].value, '1')
        at.selectbox[0].select('latin1').run()
        self.assertFalse(at.exception)
        self.assertEqual(len(at.get('download_button')), 0)

    def test_process_button_runs_current_pipeline_without_month(self):
        from streamlit.testing.v1 import AppTest
        rows = [fixtures.record(afipos='1', MES_PROCESO='09'),
                fixtures.record('PEQUENO', afipos='2', MES_PROCESO='10')]
        r34 = io.BytesIO(pd.DataFrame(rows).to_csv(index=False, sep=';').encode())
        dictionary, billing, park = fixtures.IntegrationTests().inputs()
        uploads = {'1.': r34, '2.': dictionary, '3.': billing, '4.': park}
        # AppTest has no uploader interaction API. Supply uploaded byte streams,
        # then click the real button and run the complete, unmocked processor.
        with patch('streamlit.file_uploader', side_effect=lambda label, **kwargs: uploads[label[:2]]):
            at = AppTest.from_file(str(Path(__file__).resolve().parents[1] / 'app.py'), default_timeout=20).run()
            self.assertFalse(at.button[0].disabled)
            at.button[0].click().run()
            self.assertFalse(at.exception)
            self.assertEqual(len(at.error), 0)
            self.assertEqual(len(at.get('download_button')), 1)
            result = at.session_state['result']
            self.addCleanup(result['temp'].cleanup)
            self.assertEqual(result['data'].audit.processed, 2)
            self.assertTrue(result['data'].output.exists())
