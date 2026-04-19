namespace FreeDataExportsv2;

/// <summary>
/// Discriminated union of all Excel cell value types.
///
/// Create values via the static factory methods or rely on implicit conversions:
/// <code>
/// sheet.SetCell("A1", 42);                           // Number
/// sheet.SetCell("B1", 3.14m);                        // Number (decimal)
/// sheet.SetCell("C1", "Hello");                      // Text (inline string)
/// sheet.SetCell("D1", true);                         // Boolean
/// sheet.SetCell("E1", DateTime.Today);               // Date
/// sheet.SetCell("F1", CellValue.AsFormula("SUM(A1:A5)"));
/// sheet.SetCell("G1", CellValue.AsError(ErrorCode.DivisionByZero));
/// </code>
/// </summary>
public abstract class CellValue
{
    // Prevent external subclassing while allowing inner sealed classes.
    internal CellValue() { }

    // ── Concrete subtypes ────────────────────────────────────────────────────

    /// <summary>
    /// Numeric value (integer or floating-point).
    /// Written with no <c>t</c> attribute and a <c>&lt;v&gt;</c> element.
    /// </summary>
    public sealed class Number : CellValue
    {
        public double Value { get; }
        internal Number(double value) => Value = value;
    }

    /// <summary>
    /// Plain-text string stored as an inline string (<c>t="inlineStr"</c>).
    /// Requires no shared-strings table; safe for any file size.
    /// </summary>
    public sealed class Text : CellValue
    {
        public string Value { get; }
        internal Text(string value) => Value = value;
    }

    /// <summary>
    /// Boolean value. Written as <c>t="b"</c> with <c>&lt;v&gt;1&lt;/v&gt;</c>
    /// or <c>&lt;v&gt;0&lt;/v&gt;</c>.
    /// </summary>
    public sealed class Boolean : CellValue
    {
        public bool Value { get; }
        internal Boolean(bool value) => Value = value;
    }

    /// <summary>
    /// Date/time value. Stored as an OLE Automation date number and displayed with a
    /// date number-format style that the package applies automatically.
    /// </summary>
    public sealed class Date : CellValue
    {
        public DateTime Value { get; }
        internal Date(DateTime value) => Value = value;
    }

    /// <summary>
    /// Excel formula. Optionally carries a cached result for viewers that do not
    /// recalculate on open.
    /// </summary>
    public sealed class Formula : CellValue
    {
        /// <summary>Formula text without the leading '=', e.g. <c>"SUM(A1:A10)"</c>.</summary>
        public string Expression { get; }

        /// <summary>Type of the formula's computed result, used to set the cell's <c>t</c> attribute.</summary>
        public FormulaResultType ResultType { get; }

        /// <summary>Optional cached result written to <c>&lt;v&gt;</c>. Pass null to omit.</summary>
        public object? CachedResult { get; }

        internal Formula(string expression, FormulaResultType resultType, object? cachedResult)
        {
            Expression   = expression;
            ResultType   = resultType;
            CachedResult = cachedResult;
        }
    }

    /// <summary>
    /// An Excel error cell (e.g. #DIV/0!, #N/A, #REF!).
    /// Written as <c>t="e"</c> with the error string in <c>&lt;v&gt;</c>.
    /// </summary>
    public sealed class Error : CellValue
    {
        public ErrorCode Code { get; }
        internal Error(ErrorCode code) => Code = code;
    }

    // ── Factory methods ──────────────────────────────────────────────────────

    /// <param name="value">Integer value.</param>
    public static Number  Of(int value)      => new(value);
    /// <param name="value">Long value.</param>
    public static Number  Of(long value)     => new(value);
    /// <param name="value">Float value.</param>
    public static Number  Of(float value)    => new(value);
    /// <param name="value">Double value.</param>
    public static Number  Of(double value)   => new(value);
    /// <param name="value">Decimal value (converted to double).</param>
    public static Number  Of(decimal value)  => new((double)value);
    /// <param name="value">Text string (must not be null).</param>
    public static Text    Of(string value)   => new(value ?? throw new ArgumentNullException(nameof(value)));
    /// <param name="value">Boolean value.</param>
    public static Boolean Of(bool value)     => new(value);
    /// <param name="value">Date/time value (stored as OA date).</param>
    public static Date    Of(DateTime value) => new(value);

    /// <summary>Creates a formula cell.</summary>
    /// <param name="expression">Formula text without the leading '='.</param>
    /// <param name="resultType">The type of the formula result (default: Number).</param>
    /// <param name="cachedResult">Optional cached result to write to &lt;v&gt;.</param>
    public static Formula AsFormula(string expression,
                                    FormulaResultType resultType = FormulaResultType.Number,
                                    object? cachedResult = null)
        => new(expression ?? throw new ArgumentNullException(nameof(expression)),
               resultType, cachedResult);

    /// <summary>Creates an error cell.</summary>
    public static Error AsError(ErrorCode code) => new(code);

    // ── Implicit conversions ──────────────────────────────────────────────────
    // These allow passing .NET primitives directly wherever CellValue is expected.

    public static implicit operator CellValue(int value)      => Of(value);
    public static implicit operator CellValue(long value)     => Of(value);
    public static implicit operator CellValue(float value)    => Of(value);
    public static implicit operator CellValue(double value)   => Of(value);
    public static implicit operator CellValue(decimal value)  => Of(value);
    public static implicit operator CellValue(string value)   => Of(value);
    public static implicit operator CellValue(bool value)     => Of(value);
    public static implicit operator CellValue(DateTime value) => Of(value);
}

// ── Supporting enums ─────────────────────────────────────────────────────────

/// <summary>
/// Specifies the result type of a formula, which controls the <c>t</c> attribute
/// on the cell element when a cached value is present.
/// </summary>
public enum FormulaResultType
{
    /// <summary>Numeric result — no <c>t</c> attribute (default).</summary>
    Number,
    /// <summary>Text result — <c>t="str"</c>.</summary>
    Text,
    /// <summary>Boolean result — <c>t="b"</c>.</summary>
    Boolean,
    /// <summary>Error result — <c>t="e"</c>.</summary>
    Error,
}

/// <summary>Standard Excel error codes.</summary>
public enum ErrorCode
{
    /// <summary>#DIV/0! — Division by zero.</summary>
    DivisionByZero,
    /// <summary>#N/A — Value not available.</summary>
    NotAvailable,
    /// <summary>#NAME? — Unrecognised formula name.</summary>
    InvalidName,
    /// <summary>#NULL! — Incorrect range operator or intersection.</summary>
    NullIntersection,
    /// <summary>#NUM! — Invalid numeric value.</summary>
    InvalidNumber,
    /// <summary>#REF! — Invalid cell reference.</summary>
    InvalidReference,
    /// <summary>#VALUE! — Wrong type of argument or operand.</summary>
    InvalidValue,
}

