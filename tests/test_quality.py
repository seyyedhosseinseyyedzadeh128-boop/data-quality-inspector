import os
import subprocess
import tempfile
import unittest
from unittest.mock import patch
from pathlib import Path
from openpyxl import Workbook, load_workbook
import app

# Each test gets its own automatically cleaned directory within the project.
TEMP_PARENT = Path(__file__).resolve().parent.parent / 'tests_temp'

class QualityTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        TEMP_PARENT.mkdir(exist_ok=True)
        if TEMP_PARENT.is_symlink() or TEMP_PARENT.is_junction() or not TEMP_PARENT.is_dir():
            raise RuntimeError('tests_temp must be a local directory, not a link')

    def test_csv_values_and_preservation(self):
        with tempfile.TemporaryDirectory(dir=TEMP_PARENT) as d:
            p = Path(d) / 'input.csv'
            p.write_bytes(b'\xef\xbb\xbfA,B,C,D,E\n00123,,NULL,NA,0\n00123,"",NULL,NA,0\n x , ,z,z,z\n')
            original = p.read_bytes()
            headers, rows = app.read_csv(p)
            self.assertEqual(rows[0], ['00123', None, 'NULL', 'NA', '0'])
            self.assertEqual(app.inspect(headers, rows), ([0, 2, 0, 0, 0], [1, 1, 0, 0, 0], 1))
            self.assertEqual(p.read_bytes(), original)

    def test_csv_options_and_errors(self):
        with tempfile.TemporaryDirectory(dir=TEMP_PARENT) as d:
            p = Path(d) / 'x.csv'
            p.write_bytes('a;b\n"é;z";"two\nlines"\n'.encode('latin-1'))
            self.assertEqual(app.read_csv(p, 'latin-1', ';')[1][0], ['é;z', 'two\nlines'])
            with self.assertRaises(UnicodeError): app.read_csv(p)
            p.write_text('a,b\n"unterminated', encoding='utf-8')
            with self.assertRaisesRegex(app.InspectionError, 'CSV record 2: unterminated'): app.read_csv(p)
            p.write_text('a,b\n1\n', encoding='utf-8')
            with self.assertRaises(app.InspectionError): app.read_csv(p)
            p.write_text('a,b\n1,2,3\n', encoding='utf-8')
            with self.assertRaises(app.InspectionError): app.read_csv(p)

    def test_csv_quote_validation(self):
        with tempfile.TemporaryDirectory(dir=TEMP_PARENT) as d:
            p = Path(d) / 'input.csv'
            for content, message in [
                ('a,b\nabc"def,x\n', 'quote inside unquoted field'),
                ('a,b\n"abc"x,y\n', 'invalid character after closing quote'),
                ('a,b\n"abc', 'unterminated quoted field'),
            ]:
                with self.subTest(content=content):
                    p.write_text(content, encoding='utf-8')
                    with self.assertRaisesRegex(app.InspectionError, 'CSV record 2: ' + message):
                        app.read_csv(p)
            p.write_bytes(b'a;b;c\n"001;23";"line1\nline2";"a""b"\n')
            self.assertEqual(app.read_csv(p, delimiter=';')[1], [['001;23', 'line1\nline2', 'a"b']])
            crlf = Path(d) / 'crlf.csv'
            crlf.write_bytes(b'a;b\r\n"line1\r\nline2";value\r\n')
            self.assertEqual(app.read_csv(crlf, delimiter=';')[1], [['line1\r\nline2', 'value']])

    def test_malformed_csv_does_not_create_report(self):
        with tempfile.TemporaryDirectory(dir=TEMP_PARENT) as d:
            p = Path(d) / 'input.csv'
            destination = Path(d) / 'report.txt'
            p.write_text('a,b\nabc"def,x\n', encoding='utf-8')
            with patch('builtins.input', side_effect=[str(p), '', '', str(destination)]), patch('builtins.print') as output:
                app.main()
            output.assert_called_with('Error: CSV record 2: quote inside unquoted field')
            self.assertFalse(destination.exists())

    def test_sanitized_expected_errors(self):
        with tempfile.TemporaryDirectory(dir=TEMP_PARENT) as d:
            root = Path(d)
            source = root / 'SensitiveInputXYZ.csv'
            source.write_bytes(b'a,b\n1,2\n')
            bad_csv = root / 'SensitiveMalformedXYZ.csv'
            bad_csv.write_bytes(b'a,b\nabc"def,x\n')
            corrupt = root / 'SensitiveWorkbookXYZ.xlsx'
            corrupt.write_bytes(b'not an xlsx archive')
            destination = root / 'SensitiveOutputXYZ.txt'

            def check(inputs, expected, patches=()):
                with patch('builtins.input', side_effect=inputs), patch('builtins.print') as printed:
                    from contextlib import ExitStack
                    with ExitStack() as stack:
                        for item in patches:
                            stack.enter_context(item)
                        app.main()
                message = ' '.join(str(call) for call in printed.call_args_list)
                self.assertIn(expected, message)
                for secret in ('Sensitive', str(root), 'raw-secret-details'):
                    self.assertNotIn(secret, message)

            check([str(root / 'SensitiveMissingXYZ.csv')], 'Input file does not exist')
            self.assertFalse(destination.exists())
            check([str(bad_csv), '', ''], 'quote inside unquoted field')
            self.assertFalse(destination.exists())
            check([str(corrupt)], 'XLSX file is corrupt or unsupported')
            self.assertFalse(destination.exists())
            check([str(source), '', ''], 'Permission denied',
                  [patch('app.read_csv', side_effect=PermissionError('raw-secret-details Sensitive'))])
            self.assertFalse(destination.exists())
            destination.write_bytes(b'existing')
            check([str(source), '', '', str(destination)], 'already exists')
            self.assertEqual(destination.read_bytes(), b'existing')
            check([str(source), '', '', str(destination)], 'incomplete output may remain',
                  [patch('app.write_report', side_effect=app.InspectionError(
                      'Report write or close failed; an incomplete output may remain at the destination.'))])
            self.assertEqual(destination.read_bytes(), b'existing')

    def test_unexpected_error_is_not_suppressed(self):
        with tempfile.TemporaryDirectory(dir=TEMP_PARENT) as d:
            source = Path(d) / 'input.csv'
            source.write_bytes(b'a\n1\n')
            with patch('builtins.input', side_effect=[str(source), '', '']), patch('app.read_csv', side_effect=RuntimeError('bug')):
                with self.assertRaisesRegex(RuntimeError, 'bug'):
                    app.main()

    def test_headers(self):
        for row in (['a', ''], ['a', ' '], ['a', 'a']):
            with self.assertRaises(app.InspectionError): app.headers_check(row)

    def test_types_and_whitespace(self):
        rows = [[None, ' x'], [None, ' x'], [1, 'x'], ['1', 'x'], [True, 'x'], [None, '\u2003x\u2003'], [None, 'x y']]
        self.assertEqual(app.inspect(['a', 'b'], rows), ([4, 0], [0, 3], 1))

    def test_xlsx_and_output(self):
        with tempfile.TemporaryDirectory(dir=TEMP_PARENT) as d:
            p = Path(d) / 'x.xlsx'
            b = Workbook(); s = b.active; s.title = 'First'
            s.append(['id', 'value', 'formula']); s.append(['00123', None, '=1+1'])
            s.append([123, 0, 'NA']); s['A3'].number_format = '00000'
            b.create_sheet('Second'); b.save(p); b.close()
            before = p.read_bytes()
            b = load_workbook(p, read_only=True, data_only=False)
            try:
                headers, rows = app.read_sheet(b, 'First')
                self.assertEqual(rows[0], ['00123', None, '=1+1'])
                self.assertEqual(rows[1], [123, 0, 'NA'])
                with self.assertRaises(app.InspectionError): app.read_sheet(b, 'absent')
            finally: b.close()
            self.assertEqual(p.read_bytes(), before)
            dest = Path(d) / 'report.txt'; dest.write_text('existing')
            with self.assertRaises(FileExistsError): app.write_report(p, dest, 'new')
            self.assertEqual(dest.read_text(), 'existing')
            with self.assertRaises(app.InspectionError): app.write_report(p, p.parent / '.' / p.name, 'new')

    def test_xlsx_formula_identity_and_preservation(self):
        with tempfile.TemporaryDirectory(dir=TEMP_PARENT) as d:
            p = Path(d) / 'formulas.xlsx'
            book = Workbook()
            sheet = book.active
            sheet.append(['value', 'marker'])
            sheet.append(['=1+1', 'formula'])
            sheet.append(['=1+1', 'formula'])
            sheet.append(['=1+1', 'formula'])
            sheet['A4'].data_type = 's'  # Same characters stored as literal text.
            sheet.append(['ordinary', 'text'])
            sheet.append(['ordinary', 'text'])
            sheet.append([1, 'number'])
            sheet.append(['1', 'number'])
            sheet.append([None, 'missing'])
            sheet.append([None, 'missing'])
            book.save(p)
            book.close()
            before = p.read_bytes()
            book = load_workbook(p, read_only=True, data_only=False)
            try:
                headers, rows = app.read_sheet(book, book.sheetnames[0])
            finally:
                book.close()
            self.assertEqual(rows[0][0], '=1+1')
            self.assertIsInstance(rows[0][0], app.FormulaText)
            self.assertIs(type(rows[2][0]), str)
            self.assertEqual(app.inspect(headers, rows), ([2, 0], [0, 0], 3))
            self.assertEqual(p.read_bytes(), before)

    def test_xlsx_error_identity_and_preservation(self):
        with tempfile.TemporaryDirectory(dir=TEMP_PARENT) as d:
            p = Path(d) / 'errors.xlsx'
            book = Workbook()
            sheet = book.active
            sheet.append(['value', 'marker'])
            sheet.append(['#N/A', 'error'])
            sheet.append(['#N/A', 'error'])
            sheet.append(['#N/A', 'error'])
            sheet['A4'].data_type = 's'  # Same characters stored as literal text.
            sheet.append(['#DIV/0!', 'error'])
            sheet.append(['=1+1', 'formula'])
            sheet.append(['=1+1', 'formula'])
            sheet.append(['=1+1', 'formula'])
            sheet['A8'].data_type = 's'
            sheet.append([None, 'missing'])
            sheet.append([None, 'missing'])
            book.save(p)
            book.close()
            before = p.read_bytes()
            book = load_workbook(p, read_only=True, data_only=False)
            try:
                headers, rows = app.read_sheet(book, book.sheetnames[0])
            finally:
                book.close()
            self.assertIsInstance(rows[0][0], app.ExcelErrorText)
            self.assertEqual(rows[0][0], '#N/A')
            self.assertIs(type(rows[2][0]), str)
            self.assertIsInstance(rows[3][0], app.ExcelErrorText)
            self.assertEqual(rows[3][0], '#DIV/0!')
            self.assertIsInstance(rows[4][0], app.FormulaText)
            self.assertIs(type(rows[6][0]), str)
            self.assertEqual(app.inspect(headers, rows[:4])[2], 1)
            self.assertEqual(app.inspect(headers, rows[4:7])[2], 1)
            self.assertEqual(app.inspect(headers, rows), ([2, 0], [0, 0], 3))
            self.assertEqual(p.read_bytes(), before)

    def test_report_privacy_and_settings(self):
        headers = ['ConfidentialHeaderXYZ', 'PrivateColumnABC']
        rows = [['SecretCellValueXYZ', None], ['SecretCellValueXYZ', None]]
        source = Path('private_directory') / 'SensitiveFilenameXYZ.xlsx'
        result = app.report(source, 'XLSX; worksheet=SecretWorksheetXYZ', headers, rows)
        for secret in (str(source), 'private_directory', 'SensitiveFilenameXYZ',
                       'SecretWorksheetXYZ', 'ConfidentialHeaderXYZ',
                       'PrivateColumnABC', 'SecretCellValueXYZ'):
            self.assertNotIn(secret, result)
        for expected in ('File type: XLSX', 'Rows: 2', 'Columns: 2',
                         'Missing cells: 2', 'Duplicate rows after first: 1',
                         'Cells with edge whitespace: 0', 'Column 2: missing=2'):
            self.assertIn(expected, result)
        csv_report = app.report(Path('SensitiveFilenameXYZ.csv'),
                                "CSV; encoding=utf-8-sig; delimiter=';'", headers, rows)
        self.assertIn('File type: CSV', csv_report)
        self.assertIn("encoding=utf-8-sig; delimiter=';'", csv_report)
        self.assertNotIn('SensitiveFilenameXYZ', csv_report)
        self.assertNotIn('ConfidentialHeaderXYZ', csv_report)
        self.assertNotIn('SecretCellValueXYZ', csv_report)

    def test_report_success_and_existing_destination(self):
        with tempfile.TemporaryDirectory(dir=TEMP_PARENT) as d:
            source = Path(d) / 'input.csv'
            source.write_bytes(b'original')
            dest = Path(d) / 'report.txt'
            app.write_report(source, dest, 'complete report\n')
            self.assertEqual(dest.read_bytes(), b'complete report\n')
            with self.assertRaises(FileExistsError):
                app.write_report(source, dest, 'replacement')
            self.assertEqual(dest.read_bytes(), b'complete report\n')
            self.assertEqual(source.read_bytes(), b'original')

    def test_windows_report_success_uses_utf8_lf_and_leaves_no_temporary_file(self):
        with tempfile.TemporaryDirectory(dir=TEMP_PARENT) as d:
            source = Path(d) / 'input.csv'
            source.write_bytes(b'original')
            dest = Path(d) / 'report.txt'
            with patch.object(app, 'WINDOWS', True):
                app.write_report(source, dest, 'caf\u00e9\nsecond line\n')
            self.assertEqual(dest.read_bytes(), b'caf\xc3\xa9\nsecond line\n')
            self.assertEqual(source.read_bytes(), b'original')
            self.assertEqual(set(Path(d).iterdir()), {source, dest})

    def test_windows_report_write_and_close_failures_remove_staged_file(self):
        class FailedTemporaryFile:
            def __init__(self, file, failure):
                self.file = file
                self.name = file.name
                self.failure = failure

            def __enter__(self):
                return self

            def write(self, text):
                if self.failure == 'write':
                    self.file.write(text[:3])
                    self.file.flush()
                    raise OSError('simulated partial write')
                return self.file.write(text)

            def flush(self):
                self.file.flush()

            def __exit__(self, *args):
                self.file.close()
                if self.failure == 'close':
                    raise OSError('simulated close failure')

        make_temporary_file = tempfile.NamedTemporaryFile
        with tempfile.TemporaryDirectory(dir=TEMP_PARENT) as d:
            source = Path(d) / 'input.csv'
            source.write_bytes(b'original')
            dest = Path(d) / 'report.txt'
            for failure in ('write', 'close'):
                with self.subTest(failure=failure):
                    def fail_staging(**kwargs):
                        return FailedTemporaryFile(make_temporary_file(**kwargs), failure)

                    with (
                        patch.object(app, 'WINDOWS', True),
                        patch.object(app.tempfile, 'NamedTemporaryFile', side_effect=fail_staging),
                        patch.object(app.os, 'rename') as rename,
                    ):
                        with self.assertRaisesRegex(app.InspectionError, 'failed before publication'):
                            app.write_report(source, dest, 'report\n')
                    rename.assert_not_called()
                    self.assertFalse(dest.exists())
                    self.assertEqual(set(Path(d).iterdir()), {source})
                    self.assertEqual(source.read_bytes(), b'original')

    def test_windows_report_preserves_destination_created_before_publication(self):
        real_rename = os.rename
        with tempfile.TemporaryDirectory(dir=TEMP_PARENT) as d:
            source = Path(d) / 'input.csv'
            source.write_bytes(b'original')
            dest = Path(d) / 'report.txt'

            def create_competing_destination(staged_path, destination):
                staged = Path(staged_path)
                self.assertEqual(staged.parent, dest.parent)
                self.assertEqual(staged.read_bytes(), b'complete report\n')
                self.assertEqual(destination, dest)
                with dest.open('xb') as competing:
                    competing.write(b'competing report')
                if os.name == 'nt':
                    return real_rename(staged_path, destination)
                raise FileExistsError('simulated Windows rename conflict')

            with (
                patch.object(app, 'WINDOWS', True),
                patch.object(app.os, 'rename', side_effect=create_competing_destination),
            ):
                with self.assertRaises(FileExistsError):
                    app.write_report(source, dest, 'complete report\n')
            self.assertEqual(dest.read_bytes(), b'competing report')
            self.assertEqual(source.read_bytes(), b'original')
            self.assertEqual(set(Path(d).iterdir()), {source, dest})

    def test_report_write_and_close_failures_leave_destination_untouched(self):
        class FailedOutput:
            def __init__(self, failure):
                self.failure = failure

            def __enter__(self):
                return self

            def write(self, text):
                if self.failure == 'write':
                    raise OSError('simulated write failure')

            def __exit__(self, *args):
                if self.failure == 'close':
                    raise OSError('simulated close failure')

        with tempfile.TemporaryDirectory(dir=TEMP_PARENT) as d:
            source = Path(d) / 'input.csv'
            dest = Path(d) / 'report.txt'
            for failure in ('write', 'close'):
                with self.subTest(failure=failure):
                    # Model a destination replaced by another actor after open.
                    dest.write_bytes(b'replacement owned by another actor')
                    with (
                        patch.object(app, 'WINDOWS', False),
                        patch.object(Path, 'open', return_value=FailedOutput(failure)),
                        patch.object(Path, 'unlink') as unlink,
                    ):
                        with self.assertRaisesRegex(app.InspectionError, 'incomplete output may remain'):
                            app.write_report(source, dest, 'report')
                        unlink.assert_not_called()
                    self.assertEqual(dest.read_bytes(), b'replacement owned by another actor')

@unittest.skipUnless(os.name == 'nt', 'Windows filesystem link tests require Windows')
class WindowsPathSafetyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        TEMP_PARENT.mkdir(exist_ok=True)
        if TEMP_PARENT.is_symlink() or TEMP_PARENT.is_junction() or not TEMP_PARENT.is_dir():
            raise RuntimeError('tests_temp must be a local directory, not a link')

    def check_existing_link(self, kind, target_input):
        with tempfile.TemporaryDirectory(dir=TEMP_PARENT) as d:
            root = Path(d)
            source = root / 'input.csv'
            other = root / 'other.dat'
            destination = root / 'report.txt'
            source.write_bytes(b'synthetic input')
            other.write_bytes(b'synthetic unrelated file')
            target = source if target_input else other
            try:
                if kind == 'symbolic':
                    os.symlink(target, destination)
                else:
                    os.link(target, destination)
            except (OSError, NotImplementedError) as exc:
                self.skipTest(f'{kind} link creation unavailable: {type(exc).__name__} (winerror={getattr(exc, "winerror", None)})')
            with self.assertRaises((FileExistsError, app.InspectionError)):
                app.write_report(source, destination, 'report')
            self.assertEqual(source.read_bytes(), b'synthetic input')
            self.assertEqual(other.read_bytes(), b'synthetic unrelated file')
            self.assertTrue(destination.samefile(target))
            if kind == 'symbolic':
                self.assertTrue(destination.is_symlink())

    def test_symlink_to_input(self):
        self.check_existing_link('symbolic', True)

    def test_symlink_to_unrelated_file(self):
        self.check_existing_link('symbolic', False)

    def test_hard_link_to_input(self):
        self.check_existing_link('hard', True)

    def test_hard_link_to_unrelated_file(self):
        self.check_existing_link('hard', False)

    def test_dangling_destination_symlink(self):
        with tempfile.TemporaryDirectory(dir=TEMP_PARENT) as d:
            root = Path(d)
            source, target, destination = root / 'input.csv', root / 'absent.dat', root / 'report.txt'
            source.write_bytes(b'synthetic input')
            try:
                os.symlink(target, destination)
            except (OSError, NotImplementedError) as exc:
                self.skipTest(f'symbolic link creation unavailable: {type(exc).__name__} (winerror={getattr(exc, "winerror", None)})')
            with self.assertRaises((FileExistsError, app.InspectionError)):
                app.write_report(source, destination, 'report')
            self.assertTrue(destination.is_symlink())
            self.assertFalse(target.exists())
            self.assertEqual(source.read_bytes(), b'synthetic input')

    def test_destination_through_junction(self):
        with tempfile.TemporaryDirectory(dir=TEMP_PARENT) as d:
            root = Path(d)
            real = root / 'real'
            real.mkdir()
            junction = root / 'junction'
            source = root / 'input.csv'
            source.write_bytes(b'synthetic input')
            existing = real / 'report.txt'
            existing.write_bytes(b'synthetic unrelated file')
            try:
                result = subprocess.run(['cmd.exe', '/d', '/c', 'mklink', '/J', str(junction), str(real)],
                                        capture_output=True, check=False)
            except OSError as exc:
                self.skipTest(f'junction creation unavailable: {type(exc).__name__} (winerror={getattr(exc, "winerror", None)})')
            if result.returncode != 0:
                self.skipTest(f'junction creation unavailable: mklink exit code {result.returncode}')
            self.assertTrue(junction.is_junction())
            with self.assertRaises((FileExistsError, app.InspectionError)):
                app.write_report(source, junction / 'report.txt', 'report')
            self.assertEqual(source.read_bytes(), b'synthetic input')
            self.assertEqual(existing.read_bytes(), b'synthetic unrelated file')
            self.assertTrue((junction / 'report.txt').samefile(existing))

    def test_same_resolved_path_and_existing_file(self):
        with tempfile.TemporaryDirectory(dir=TEMP_PARENT) as d:
            root = Path(d)
            source = root / 'input.txt'
            source.write_bytes(b'synthetic input')
            with self.assertRaises(app.InspectionError):
                app.write_report(source, root / '.' / 'input.txt', 'report')
            destination = root / 'report.txt'
            destination.write_bytes(b'synthetic unrelated file')
            with self.assertRaises(FileExistsError):
                app.write_report(source, destination, 'report')
            self.assertEqual(source.read_bytes(), b'synthetic input')
            self.assertEqual(destination.read_bytes(), b'synthetic unrelated file')

if __name__ == '__main__': unittest.main()
