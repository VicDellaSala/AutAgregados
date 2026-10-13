import io
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch
import pandas as pd
import test_processing as fixtures
from agregados.archivos import open_r34
from agregados.procesamiento import process


def archive(entries):
    buffer = io.BytesIO()
    buffer.name = 'R34.ZIP'
    with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in entries:
            z.writestr(name, data)
    buffer.seek(0)
    return buffer


class ZipTests(unittest.TestCase):
    def test_zip_matches_csv_with_other_files_and_nested_folder(self):
        rows = [fixtures.record(), fixtures.record('CARROPAGO'), fixtures.record('PEQUENO')]
        raw = pd.DataFrame(rows).to_csv(index=False, sep=';').encode('utf-8')
        zipped = archive([('ibo/RNC.CSV', raw), ('ibo/Referencia.xlsx', b'ignored')])
        with tempfile.TemporaryDirectory() as temp:
            plain = fixtures.IntegrationTests().run_process(rows, Path(temp)/'plain.xlsx')
            packed = process(zipped, *fixtures.IntegrationTests().inputs(), Path(temp)/'zip.xlsx', chunksize=1)
            pd.testing.assert_frame_equal(plain.preview, packed.preview)
            pd.testing.assert_frame_equal(plain.summary, packed.summary)
            self.assertEqual(plain.audit.read, packed.audit.read)
            self.assertFalse(zipped.closed)
            self.assertEqual(sorted(p.name for p in Path(temp).iterdir()), ['plain.xlsx', 'zip.xlsx'])

    def test_disk_zip_and_no_extraction(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp)/'source.zip'
            path.write_bytes(archive([('../R34.csv', b'header\nvalue\n')]).getvalue())
            with open_r34(path) as csv:
                self.assertIsInstance(csv, zipfile.ZipExtFile)
                self.assertEqual(csv.readline(), b'header\n')
                self.assertEqual(csv.read(), b'value\n')
            self.assertEqual(list(Path(temp).iterdir()), [path])

    def test_reject_zero_or_multiple_csvs(self):
        for entries in [[('readme.txt', b'none')], [('a.csv', b'a'), ('b.csv', b'b')]]:
            with self.assertRaisesRegex(ValueError, 'único CSV'):
                with open_r34(archive(entries)):
                    pass

    def test_reject_corrupt_zip(self):
        source = io.BytesIO(b'not a zip')
        source.name = 'r34.zip'
        with self.assertRaisesRegex(ValueError, 'dañado'):
            with open_r34(source):
                pass

    def test_reject_uncompressed_limit_before_reading(self):
        with patch('agregados.archivos.MAX_UNCOMPRESSED_BYTES', 10):
            with self.assertRaisesRegex(ValueError, 'descomprimidos'):
                with open_r34(archive([('r34.csv', b'X'*11)])):
                    pass

    def test_encrypted_flag(self):
        data = bytearray(archive([('r34.csv', b'a')]).getvalue())
        # Flag bit 0 in both ZIP headers marks an encrypted member.
        for signature, offset in [(b'PK\x03\x04', 6), (b'PK\x01\x02', 8)]:
            pos = data.index(signature) + offset
            data[pos] |= 1
        source = io.BytesIO(data)
        source.name = 'r34.zip'
        with self.assertRaisesRegex(ValueError, 'contraseña'):
            with open_r34(source):
                pass

    def test_stream_does_not_read_entire_member(self):
        source = archive([('r34.csv', b'header\n'+b'12345678\n'*200000)])
        with patch.object(zipfile.ZipFile, 'read', side_effect=AssertionError('No complete archive reads')):
            with open_r34(source) as csv:
                self.assertEqual(csv.read(7), b'header\n')
                self.assertLess(csv.tell(), 1800007)

    def test_csv_header_utf8_boundary(self):
        # Only the header is decoded for sniffing, not a fixed-size prefix
        # potentially ending halfway through an accented character in a row.
        row = fixtures.record(TERMINAL='Á'*20000)
        with tempfile.TemporaryDirectory() as temp:
            result = fixtures.IntegrationTests().run_process([row], Path(temp)/'out.xlsx')
            self.assertEqual(result.audit.processed, 1)
