# xlsx-formula-engine

A minimal zero-dependency Kotlin library for evaluating formulas from `.xlsx` files without Excel, LibreOffice, Apache POI, or any external runtime dependencies.

Disclaimer: this engine is not a full Excel-compatible runtime; verify results before using them in production or financial decisions. The author disclaims any responsibility for damages or losses resulting from its use.

## Scope and Limitations

- `.xlsx` files are read directly through `ZIP + XML (StAX)`.
- Supported formulas:
  - Arithmetic: `+`, `-`, `*`, `/`, `^`
  - Parentheses
  - Cell references: `A1`, `$B$2`
  - Cross-sheet references: `Sheet2!C3`
  - Ranges: `A1:B10`, `Sheet1!A1:B2`
  - Functions: `IF`, `IFERROR`, `IFNA`, `INDEX`, `MATCH`, `VLOOKUP`, `HLOOKUP`, `OFFSET`, `INDIRECT`, `CHOOSE`, `SUM`, `SUMIF`, `SUMIFS`, `AREAS`, `AVEDEV`, `CORREL`, `GEOMEAN`, `HARMEAN`, `MAXA`, `MODE`, `COUNTA`, `COUNTBLANK`, `COUNTIF`, `COUNTIFS`, `AVERAGEIF`, `AVERAGEIFS`, `RANK`, `PERCENTILE`, `QUARTILE`, `STDEV`, `STDEVP`, `VAR`, `VARP`, `SUMX2MY2`, `SUMX2PY2`, `SUMXMY2`, `MEDIAN`, `SMALL`, `LARGE`, `PRODUCT`, `SUMSQ`, `CEILING`, `FLOOR`, `MROUND`, `SIGN`, `RAND`, `TODAY`, `NOW`, `PI`, `SUMPRODUCT`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `AND`, `OR`, `NOT`, `ISERROR`, `ISERR`, `ISNA`, `ISNUMBER`, `ISBLANK`, `ISLOGICAL`, `ISNONTEXT`, `ISTEXT`, `ISEVEN`, `ISODD`, `TRUE`, `FALSE`, `ABS`, `ACOS`, `ACOSH`, `ASIN`, `ASINH`, `ATAN2`, `ATANH`, `COSH`, `RADIANS`, `SIN`, `SINH`, `TAN`, `TANH`, `SQRTPI`, `TRUNC`, `QUOTIENT`, `COLUMN`, `COLUMNS`, `ROWS`, `POWER`, `INT`, `MOD`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `LN`, `SQRT`, `EXP`, `NORMSDIST`, `LEN`, `LEFT`, `LEFTB`, `RIGHT`, `MID`, `MIDB`, `LENB`, `CONCATENATE`, `UPPER`, `LOWER`, `TRIM`, `REPT`, `EXACT`, `VALUE`, `CHAR`, `CLEAN`, `CODE`, `PROPER`, `REPLACE`, `TEXT`, `DOLLAR`, `FIXED`, `FIND`, `SEARCH`, `SUBSTITUTE`, `TIME`, `TIMEVALUE`, `DATE`, `DATEDIF`, `DAYS360`, `EDATE`, `EOMONTH`, `NETWORKDAYS`, `YEAR`, `MONTH`, `HOUR`, `SECOND`, `DAY`, `LOOKUP`, `LOG10`, `EFFECT`, `EXPONDIST`, `COMBIN`, `GCD`, `GESTEP`, `PMT`, `PV`, `FV`, `NPER`, `RATE`, `NPV`, `IRR`, `PPMT`
- Supported cell types: numeric, shared string, inline string, bool, error.

## Not Supported

- Full compatibility with all Excel functions
- Named ranges
- Array formulas
- Shared formulas (`t="shared"` with formula propagation)
- Date/time formats and styles

## Usage

```kotlin
import com.rinat.xlsxengine.XlsxFormulaEngine
import java.nio.file.Paths

val engine = XlsxFormulaEngine.fromFile(Paths.get("report.xlsx"))

val a1 = engine.evaluateCell("Sheet1", "A1")
val row2 = engine.evaluateRow("Sheet1", 2)
val colA = engine.evaluateColumn("Sheet1", "A")
val all = engine.evaluateAll()
```

## Adding Functions

To add another Excel-style function, implement `FormulaFunctionHandler` and register it in `FunctionRegistry`.

Example: add `DOUBLE(value)`, which returns the numeric argument multiplied by two.

```kotlin
internal class DoubleFunction : FormulaFunctionHandler {
    override fun evaluate(
        context: FunctionEvaluationContext,
        currentSheet: String,
        args: List<FormulaNode>
    ): EvalValue {
        if (args.size != 1) {
            return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        }

        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))

        return EvalValue.Scalar(CellValue.NumberValue(number * 2.0))
    }
}
```

Then add it to `FunctionRegistry`:

```kotlin
"DOUBLE" to DoubleFunction()
```

After registering it, formulas such as `=DOUBLE(A1)` can be parsed and evaluated. Add regression tests for new functions and update the supported-function list in this README.

## Build

```bash
./gradlew test
```

## Architecture

- `XlsxReader`:
  - parses the workbook/sheets/rels/sharedStrings/worksheet structure
- `FormulaTokenizer` + `FormulaParser`:
  - build the formula AST
- `FormulaEvaluator`:
  - evaluates the AST, resolves references and ranges, and caches values
  - detects reference cycles (`#CYCLE!`)
- `XlsxFormulaEngine`:
  - public API for evaluating cells, rows, columns, and the entire workbook
