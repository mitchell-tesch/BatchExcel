namespace BatchExcel.Models;

/// <summary>An Excel cell error value such as <c>#DIV/0!</c>.</summary>
public sealed record ExcelError(string Text)
{
    private static readonly Dictionary<int, string> ByCode = new()
    {
        [2000] = "#NULL!",
        [2007] = "#DIV/0!",
        [2015] = "#VALUE!",
        [2023] = "#REF!",
        [2029] = "#NAME?",
        [2036] = "#NUM!",
        [2042] = "#N/A",
        [2043] = "#GETTING_DATA",
        [2045] = "#SPILL!",
        [2046] = "#CONNECT!",
        [2047] = "#BLOCKED!",
        [2048] = "#UNKNOWN!",
        [2049] = "#FIELD!",
        [2050] = "#CALC!",
    };

    /// <summary>True for the errors the OpenXML file format can store as a typed error cell.</summary>
    public bool IsStandard => Text is "#NULL!" or "#DIV/0!" or "#VALUE!" or "#REF!" or "#NAME?" or "#NUM!" or "#N/A" or "#GETTING_DATA";

    public override string ToString() => Text;

    /// <summary>
    /// Converts a cell error read via COM (VT_ERROR, surfaced as int 0x800A0000 + error code)
    /// to an <see cref="ExcelError"/>. Any other value is returned unchanged.
    /// </summary>
    public static object? FromComValue(object? value)
    {
        if (value is not int hr || (hr & unchecked((int)0xFFFF0000)) != unchecked((int)0x800A0000))
            return value;

        var code = hr & 0xFFFF;
        return new ExcelError(ByCode.TryGetValue(code, out var text) ? text : $"#ERROR({code})");
    }
}
