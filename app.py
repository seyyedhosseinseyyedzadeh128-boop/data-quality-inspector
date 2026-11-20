import csv
import codecs
import os
import tempfile
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

WINDOWS = os.name == 'nt'

class InspectionError(Exception):
    pass

class FormulaText(str):
    """Formula text tagged for type-sensitive duplicate comparison."""

class ExcelErrorText(str):
    """Excel error text tagged for type-sensitive duplicate comparison."""

def headers_check(row):
    seen = {}
    for i, value in enumerate(row, 1):
        if value is None or not str(value).strip():
            raise InspectionError(f'Empty header at column {i}')
        name = str(value)
        if name in seen:
            raise InspectionError(f'Duplicate headers at columns {seen[name]} and {i}')
        seen[name] = i
    if not row:
        raise InspectionError('No headers found')
    return [str(v) for v in row]

def validated_csv_lines(source, delimiter):
    """Validate quote placement while streaming physical lines to csv.reader."""
    state = 'start'
    record = 1
    for line in source:
        i = 0
        while i < len(line):
            char = line[i]
            if state == 'quoted':
                if char == '"':
                    state = 'closed'
            elif char in '\r\n':
                if char == '\r' and i + 1 < len(line) and line[i + 1] == '\n':
                    i += 1
                record += 1
                state = 'start'
            elif state == 'closed':
                if char == '"':
                    state = 'quoted'
                elif char == delimiter:
                    state = 'start'
                else:
                    raise InspectionError(f'CSV record {record}: invalid character after closing quote')
            elif char == delimiter:
                state = 'start'
            elif char == '"':
                if state != 'start':
                    raise InspectionError(f'CSV record {record}: quote inside unquoted field')
                state = 'quoted'
            else:
                state = 'plain'
            i += 1
        yield line
    if state == 'quoted':
        raise InspectionError(f'CSV record {record}: unterminated quoted field')

def read_csv(path, encoding='utf-8-sig', delimiter=','):
    if len(delimiter) != 1 or delimiter in '\r\n"':
        raise InspectionError('Delimiter must be one non-quote, non-newline character')
    with open(path, encoding=encoding, newline='') as f:
        reader = csv.reader(validated_csv_lines(f, delimiter), delimiter=delimiter, strict=True)
        headers = headers_check(next(reader, []))
        rows = []
        for number, row in enumerate(reader, 2):
            if len(row) != len(headers):
                raise InspectionError(f'CSV record {number}: expected {len(headers)} fields, found {len(row)}')
            rows.append([None if v == '' else v for v in row])
    return headers, rows

def read_sheet(book, name):
    if name not in book.sheetnames:
        raise InspectionError('Invalid worksheet selection')
    rows = book[name].iter_rows()
    headers = headers_check([cell.value for cell in next(rows, [])])
    result = []
    for row in rows:
        values = []
        for cell in row:
            value = cell.value
            if cell.data_type == 'f':
                if not isinstance(value, str):
                    raise InspectionError('Unsupported formula representation in worksheet')
                value = FormulaText(value)
            elif cell.data_type == 'e' and value is not None:
                if not isinstance(value, str):
                    raise InspectionError('Unsupported error representation in worksheet')
                value = ExcelErrorText(value)
            values.append(None if value == '' else value)
        result.append(values)
    return headers, result

def inspect(headers, rows):
    missing = [0] * len(headers)
    spaces = [0] * len(headers)
    seen = set()
    duplicates = 0
    for row in rows:
        key = tuple(('missing',) if v is None else (type(v).__name__, v) for v in row)
        if key in seen:
            duplicates += 1
        seen.add(key)
        for i, value in enumerate(row):
            if value is None:
                missing[i] += 1
            elif isinstance(value, str) and value and (value[0].isspace() or value[-1].isspace()):
                spaces[i] += 1
    return missing, spaces, duplicates

def report(path, settings, headers, rows):
    missing, spaces, duplicates = inspect(headers, rows)
    file_type = Path(path).suffix.lower()
    safe_settings = 'XLSX' if file_type == '.xlsx' else settings
    lines = ['DATA QUALITY REPORT', f'File type: {file_type.lstrip(".").upper()}', f'Settings: {safe_settings}',
             f'Rows: {len(rows)}', f'Columns: {len(headers)}', f'Missing cells: {sum(missing)}',
             f'Duplicate rows after first: {duplicates}', f'Cells with edge whitespace: {sum(spaces)}',
             'Column counts (column names omitted):']
    lines += [f'Column {i}: missing={m}, whitespace={s}' for i, (m, s) in enumerate(zip(missing, spaces), 1)]
    lines += ['', 'Empty fields/cells are missing; NULL, NA, 0 and whitespace-only text are not.',
              'Duplicate comparison uses exact parsed types and values; missing values compare equal.',
              'Formula text is not evaluated. Excel display formatting is not preserved.',
              'Worksheet used dimensions define the table. The entire table is loaded into memory.']
    return '\n'.join(lines) + '\n'

def write_report(source, destination, text):
    source, destination = Path(source), Path(destination)
    if source.resolve() == destination.resolve():
        raise InspectionError('Report path must differ from input')
    if destination.suffix.lower() != '.txt':
        raise InspectionError('Report path must end in .txt')
    if WINDOWS:
        publish_windows_report(destination, text)
        return
    # Open separately so an existing destination retains its original error.
    f = destination.open('x', encoding='utf-8', newline='\n')
    try:
        with f:
            f.write(text)
    except (OSError, UnicodeError) as exc:
        # A pathname identity check followed by unlink is racy: another file
        # could replace it between those operations. Leave it untouched.
        raise InspectionError('Report write or close failed; an incomplete output may remain at the destination.') from exc

def publish_windows_report(destination, text):
    temporary = tempfile.NamedTemporaryFile(
        mode='w', encoding='utf-8', newline='\n', dir=destination.parent, delete=False
    )
    try:
        try:
            with temporary:
                temporary.write(text)
                temporary.flush()
        except (OSError, UnicodeError) as exc:
            raise InspectionError('Report write or close failed before publication.') from exc
        os.rename(temporary.name, destination)
    except BaseException as exc:
        try:
            Path(temporary.name).unlink(missing_ok=True)
        except OSError as cleanup_error:
            exc.add_note(f'Temporary report cleanup failed: {cleanup_error}')
        raise

def main():
    try:
        path = Path(input('Input CSV or XLSX path: ').strip().strip('"'))
        if not path.is_file():
            raise InspectionError('Input file does not exist')
        if path.suffix.lower() == '.csv':
            encoding = input('CSV encoding [utf-8-sig]: ').strip() or 'utf-8-sig'
            delimiter = input('CSV delimiter [,]: ') or ','
            try:
                codecs.lookup(encoding)
            except LookupError:
                raise InspectionError('Unknown CSV encoding; choose a supported encoding.') from None
            headers, rows = read_csv(path, encoding, delimiter)
            settings = f'CSV; encoding={encoding}; delimiter={delimiter!r}'
        elif path.suffix.lower() == '.xlsx':
            book = load_workbook(path, read_only=True, data_only=False, keep_links=False)
            try:
                for i, name in enumerate(book.sheetnames, 1):
                    print(f'{i}. {name}')
                selection = input('Worksheet number [1]: ').strip() or '1'
                if not selection.isdigit() or not 1 <= int(selection) <= len(book.sheetnames):
                    raise InspectionError('Invalid worksheet selection')
                name = book.sheetnames[int(selection)-1]
                headers, rows = read_sheet(book, name)
                settings = 'XLSX'
            finally:
                book.close()
        else:
            raise InspectionError('Unsupported format; choose CSV or XLSX')
        output = input('New report TXT path: ').strip().strip('"')
        if not output:
            raise InspectionError('Report path is required')
        write_report(path, output, report(path, settings, headers, rows))
        print('Report created successfully.')
    except InspectionError as exc:
        print(f'Error: {exc}')
    except FileExistsError:
        print('Error: Report destination already exists; choose a new path.')
    except PermissionError:
        print('Error: Permission denied; check access to the input and report destination.')
    except FileNotFoundError:
        print('Error: Input file or report directory was not found; check the paths.')
    except (UnicodeError, csv.Error):
        print('Error: CSV could not be decoded or parsed; check its encoding and format.')
    except (BadZipFile, InvalidFileException, ParseError):
        print('Error: XLSX file is corrupt or unsupported; choose a valid workbook.')
    except OSError:
        print('Error: A file operation failed; check access and available storage.')

if __name__ == '__main__':
    main()
