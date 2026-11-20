# Excel/CSV Data Quality Inspector

## Business problem

Small businesses often rely on spreadsheets for order tracking and reporting. Missing entries, repeated rows, and stray spaces can make those records harder to review and use consistently.

## Solution overview

This Python prototype inspects one local CSV file or one worksheet in an XLSX workbook and produces a separate text report. It identifies missing cells, exact duplicate rows, and leading or trailing whitespace. It runs offline and does not automatically clean records or modify the input file.

## Demonstration methodology

A fictional order workbook was prepared with known data-quality issues. The inspector was run once on the workbook, and its report was compared with the independently known fixture counts. The workbook's SHA256 hash was compared before and after the run to check that its bytes remained unchanged. The demonstration completed successfully with exit code 0.

## Verified results

| Measure | Expected | Inspector report |
|---|---:|---:|
| Synthetic order records | 100 | 100 |
| Columns | 7 | 7 |
| Missing cells | 5 | 5 |
| Duplicate rows beyond their first occurrences | 3 | 3 |
| Cells with leading or trailing whitespace | 4 | 4 |

All reported counts matched the expected fixture counts. The input workbook's SHA256 hash was unchanged.

## Data privacy considerations

The demonstration used entirely fictional data. The report omitted raw cell values, worksheet names, and input paths, and identified columns by position. Aggregate statistics can still reveal information about a dataset. Worksheet names appear during interactive selection, and unexpected errors may expose diagnostic details.

## Current limitations

This is a prototype, not a production-ready service or a comprehensively security-validated system. It reads a selected table into memory, does not evaluate Excel formulas, and does not preserve Excel display formatting when it differs from stored values. Macro-enabled workbooks are unsupported. On Windows, the application writes the report to a temporary file, flushes and closes it, and then publishes it. A staging write or close failure leaves no partial report at the final destination. In exceptional circumstances, temporary-file cleanup may fail. On POSIX, the application writes directly to the destination using exclusive file creation, so a write failure may leave a partial report. The demonstration validates only the described synthetic example.

## Potential business applications

The inspector could support an initial review of order exports, inventory lists, or customer records before a person decides how to investigate and correct issues. Its report provides a starting point for discussion; it does not guarantee data accuracy or business outcomes.