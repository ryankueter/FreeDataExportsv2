namespace FreeDataExportsv2;

/// <summary>
/// Utilities for Excel A1-style cell and column references.
/// All row and column indices are 1-based.
/// </summary>
public static class CellReference
{
    /// <summary>
    /// Converts 1-based row and column indices to an A1-style reference.
    /// e.g. (1, 3) → "C1"
    /// </summary>
    public static string FromRowCol(int row, int col) => ColumnName(col) + row;

    /// <summary>
    /// Converts a 1-based column index to its letter name.
    /// e.g. 1 → "A", 26 → "Z", 27 → "AA"
    /// </summary>
    public static string ColumnName(int col)
    {
        var name = string.Empty;
        while (col > 0)
        {
            col--;
            name = (char)('A' + col % 26) + name;
            col /= 26;
        }
        return name;
    }

    /// <summary>
    /// Converts a column letter name to its 1-based index.
    /// e.g. "A" → 1, "Z" → 26, "AA" → 27
    /// </summary>
    public static int ColumnIndex(string colName)
    {
        int result = 0;
        foreach (char c in colName)
            result = result * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
        return result;
    }

    /// <summary>
    /// Parses an A1-style reference into its 1-based row and column indices.
    /// e.g. "C5" → (row: 5, col: 3)
    /// </summary>
    public static (int row, int col) Parse(string cellRef)
    {
        if (string.IsNullOrEmpty(cellRef))
            throw new ArgumentException("XlsxCell reference must not be empty.", nameof(cellRef));

        int i = 0;
        while (i < cellRef.Length && char.IsLetter(cellRef[i])) i++;

        if (i == 0 || i == cellRef.Length)
            throw new FormatException($"'{cellRef}' is not a valid A1-style cell reference.");

        return (int.Parse(cellRef[i..]), ColumnIndex(cellRef[..i]));
    }
}
