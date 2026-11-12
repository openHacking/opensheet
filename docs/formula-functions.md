# Formula function reference

Generated from the runtime function registry by `pnpm formula:docs` after building packages. Do not edit this table manually. See [semantics and limitations](formulas.md) before assuming Excel compatibility.

| Function | Signature | Arguments | Evaluation | Description |
| --- | --- | --- | --- | --- |
| ROUNDUP | `ROUNDUP(number, digits)` | 2–2 | scalar | Round away from zero on the shortest decimal representation; digits truncate toward zero. |
| ROUNDDOWN | `ROUNDDOWN(number, digits)` | 2–2 | scalar | Round toward zero on the shortest decimal representation; digits truncate toward zero. |
| INT | `INT(number)` | 1–1 | scalar | Floor to the nearest integer, including negative numbers. |
| MOD | `MOD(number, divisor)` | 2–2 | scalar | Remainder with divisor sign; zero divisor returns DIV/0. |
| IFNA | `IFNA(value, fallback)` | 2–2 | lazy | Evaluate fallback only for the NA error. |
| IFS | `IFS(test, value, ...)` | 2–254 | lazy | Evaluate pairs in order and return the first matching scalar value; no match returns NA. |
| DATE | `DATE(year, month, day)` | 3–3 | scalar | Construct a serial in the workbook date system, normalizing month/day overflow. |
| YEAR | `YEAR(serial)` | 1–1 | scalar | Gregorian year in the workbook date system. |
| MONTH | `MONTH(serial)` | 1–1 | scalar | Month number in the workbook date system. |
| DAY | `DAY(serial)` | 1–1 | scalar | Day of month, including the 1900 compatibility leap day. |
| DAYS | `DAYS(end, start)` | 2–2 | scalar | Subtract date serials; ISO date text is supported, ambiguous locale text is not. |
| SUMIF | `SUMIF(range, criterion, [sumRange])` | 2–3 | conditional | Sum matching numeric cells; optional value range extends from its top-left. |
| SUMIFS | `SUMIFS(sumRange, criteriaRange, criterion, ...)` | 3–255 | conditional | Sum numeric cells matching every criterion; shapes must agree. |
| COUNTIF | `COUNTIF(range, criterion)` | 2–2 | conditional | Count matching cells with case-insensitive criteria and escaped wildcards. |
| COUNTIFS | `COUNTIFS(criteriaRange, criterion, ...)` | 2–254 | conditional | Count cells matching every criterion; shapes must agree. |
| AVERAGEIF | `AVERAGEIF(range, criterion, [averageRange])` | 2–3 | conditional | Average matching numeric cells; no numeric matches returns DIV/0. |
| AVERAGEIFS | `AVERAGEIFS(averageRange, criteriaRange, criterion, ...)` | 3–255 | conditional | Average numeric cells matching every criterion; shapes must agree. |
| INDEX | `INDEX(range, row, [column])` | 2–3 | lookup | Return one cell; positive scalar indices only. |
| MATCH | `MATCH(value, vector, [matchType])` | 2–3 | lookup | Exact/wildcard or ordered approximate match; default match type is 1. |
| VLOOKUP | `VLOOKUP(value, table, column, [approximate])` | 3–4 | lookup | Exact/wildcard or ordered approximate match; approximate is the default. |
| XLOOKUP | `XLOOKUP(value, vector, results, [missing], [matchMode], [searchMode])` | 3–6 | lookup | Scalar result; exact/nearest/wildcard matching and linear or sorted binary search. |
| TRUE | `TRUE()` | 0–0 | scalar | Return the logical value true. |
| FALSE | `FALSE()` | 0–0 | scalar | Return the logical value false. |
| LEFT | `LEFT(text, [count])` | 1–2 | scalar | Leading Unicode code points; count defaults to 1. |
| RIGHT | `RIGHT(text, [count])` | 1–2 | scalar | Trailing Unicode code points; count defaults to 1. |
| MID | `MID(text, start, count)` | 3–3 | scalar | Unicode code-point substring from a 1-based position. |
| LEN | `LEN(text)` | 1–1 | scalar | Count Unicode code points, not grapheme clusters. |
| TRIM | `TRIM(text)` | 1–1 | scalar | Remove leading/trailing ASCII spaces and collapse repeated ASCII spaces. |
| UPPER | `UPPER(text)` | 1–1 | scalar | Locale-independent Unicode uppercase. |
| LOWER | `LOWER(text)` | 1–1 | scalar | Locale-independent Unicode lowercase. |
| CONCAT | `CONCAT(text, ...)` | 1–253 | text | Concatenate scalars/ranges in row-major order; 32767 UTF-16 code-unit output limit. |
| TEXTJOIN | `TEXTJOIN(delimiter, ignoreEmpty, text, ...)` | 3–254 | text | Join scalar/range text with a scalar delimiter; at most 252 text arguments. |
| SUM | `SUM(value, ...)` | 1–255 | aggregate | Sum numeric values. |
| AVERAGE | `AVERAGE(value, ...)` | 1–255 | aggregate | Mean of numeric values. |
| MIN | `MIN(value, ...)` | 1–255 | aggregate | Smallest numeric value, or zero. |
| MAX | `MAX(value, ...)` | 1–255 | aggregate | Largest numeric value, or zero. |
| COUNT | `COUNT(value, ...)` | 1–255 | aggregate | Count numbers; direct logical and numeric text arguments also count. |
| COUNTA | `COUNTA(value, ...)` | 1–255 | aggregate | Count nonblank values, including empty text and errors. |
| IF | `IF(test, whenTrue, [whenFalse])` | 2–3 | lazy | Evaluate only the selected branch; omitted false branch returns false. |
| IFERROR | `IFERROR(value, fallback)` | 2–2 | lazy | Evaluate fallback only for an error. |
| AND | `AND(value, ...)` | 1–255 | aggregate | All logical values must be true; referenced text and blanks are ignored. |
| OR | `OR(value, ...)` | 1–255 | aggregate | At least one logical value must be true; referenced text and blanks are ignored. |
| NOT | `NOT(value)` | 1–1 | scalar | Invert a logical value. |
| ROUND | `ROUND(number, digits)` | 2–2 | scalar | Round halfway away from zero; digits truncate toward zero. |
| ABS | `ABS(number)` | 1–1 | scalar | Absolute numeric value. |
