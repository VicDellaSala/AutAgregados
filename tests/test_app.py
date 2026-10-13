"""Optional Streamlit smoke test; runs when requirements.txt is installed."""
import importlib.util
import tempfile
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
