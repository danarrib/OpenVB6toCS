namespace VB6toCS.Core.CodeGeneration;

/// <summary>
/// Maps VB6 built-in function names to C# equivalents.
/// Functions listed here are emitted as transformed calls; unlisted ones are
/// emitted as-is with a // REVIEW: comment.
/// </summary>
internal static class BuiltInMap
{
    // Maps VB6 function name → a delegate that takes the raw argument list string
    // and returns the C# expression string.
    private static readonly Dictionary<string, Func<string[], string>> Functions =
        new(StringComparer.OrdinalIgnoreCase)
        {
            // String functions
            ["Len"]       = a => $"{One(a)}.Length",
            ["UCase"]     = a => $"{One(a)}.ToUpper()",
            ["LCase"]     = a => $"{One(a)}.ToLower()",
            ["Trim"]      = a => $"{One(a)}.Trim()",
            ["LTrim"]     = a => $"{One(a)}.TrimStart()",
            ["RTrim"]     = a => $"{One(a)}.TrimEnd()",
            ["Left"]      = a => $"{Arg(a,0)}.Substring(0, {Arg(a,1)})",
            ["Right"]     = a => $"{Arg(a,0)}.Substring({Arg(a,0)}.Length - {Arg(a,1)})",
            ["Mid"]       = a => a.Length >= 3
                                 ? $"{Arg(a,0)}.Substring({Arg(a,1)} - 1, {Arg(a,2)})"
                                 : $"{Arg(a,0)}.Substring({Arg(a,1)} - 1)",
            ["InStr"]     = a => a.Length >= 3
                                 ? $"({Arg(a,1)}.IndexOf({Arg(a,2)}) + 1)"
                                 : $"({Arg(a,0)}.IndexOf({Arg(a,1)}) + 1)",
            ["InStrRev"]  = a => $"({Arg(a,0)}.LastIndexOf({Arg(a,1)}) + 1)",
            ["Replace"]   = a => $"{Arg(a,0)}.Replace({Arg(a,1)}, {Arg(a,2)})",
            ["Split"]     = a => $"{Arg(a,0)}.Split({Arg(a,1)})",
            ["Join"]      = a => $"string.Join({Arg(a,1)}, {Arg(a,0)})",
            ["Space"]     = a => $"new string(' ', {One(a)})",
            ["String"]    = a =>
            {
                string charArg = Arg(a, 1);
                // VB6 String(n, c): c may be a string literal "x" → char literal 'x',
                // or a number (char code) → (char)n, or a variable → var[0].
                if (charArg.StartsWith('"') && charArg.EndsWith('"') && charArg.Length >= 3)
                    charArg = $"'{charArg[1]}'"; // "0" → '0'
                else if (long.TryParse(charArg, out _))
                    charArg = $"(char){charArg}"; // 48 → (char)48
                else
                    charArg = $"{charArg}[0]";   // variable → variable[0]
                return $"new string({charArg}, {Arg(a, 0)})";
            },
            ["StrReverse"] = a => $"new string({One(a)}.Reverse().ToArray())",
            ["Chr"]       = a => $"((char){One(a)}).ToString()",
            ["Asc"]       = a => $"(int){One(a)}[0]",
            ["Hex"]       = a => $"{One(a)}.ToString(\"X\")",
            ["Oct"]       = a => $"Convert.ToString({One(a)}, 8)",
            ["Format"]    = a => $"{Arg(a,0)}.ToString({(a.Length > 1 ? Arg(a,1) : "")}) /* Format */",

            // Type conversion
            ["CStr"]      = a => $"{One(a)}.ToString()",
            ["CInt"]      = a => $"Convert.ToInt32({One(a)})",
            ["CLng"]      = a => $"Convert.ToInt32({One(a)})",
            ["CDbl"]      = a => $"Convert.ToDouble({One(a)})",
            ["CSng"]      = a => $"Convert.ToSingle({One(a)})",
            ["CBool"]     = a => $"Convert.ToBoolean({One(a)})",
            ["CByte"]     = a => $"Convert.ToByte({One(a)})",
            ["CDate"]     = a => $"Convert.ToDateTime({One(a)})",
            ["CVDate"]    = a => $"Convert.ToDateTime({One(a)})",
            ["Val"]       = a => $"Convert.ToDouble({One(a)})",
            ["Str"]       = a => $"{One(a)}.ToString()",

            // Math functions
            ["Abs"]       = a => $"Math.Abs({One(a)})",
            ["Sqr"]       = a => $"Math.Sqrt({One(a)})",
            ["Int"]       = a => $"(int)Math.Floor({One(a)})",
            ["Fix"]       = a => $"(int)Math.Truncate({One(a)})",
            ["Sgn"]       = a => $"Math.Sign({One(a)})",
            ["Exp"]       = a => $"Math.Exp({One(a)})",
            ["Log"]       = a => $"Math.Log({One(a)})",
            ["Sin"]       = a => $"Math.Sin({One(a)})",
            ["Cos"]       = a => $"Math.Cos({One(a)})",
            ["Tan"]       = a => $"Math.Tan({One(a)})",
            ["Atn"]       = a => $"Math.Atan({One(a)})",
            ["Round"]     = a => a.Length > 1
                                 ? $"Math.Round({Arg(a,0)}, {Arg(a,1)})"
                                 : $"Math.Round({One(a)})",

            // Date/Time
            ["Now"]       = _ => "DateTime.Now",
            ["Date"]      = _ => "DateTime.Today",
            ["Time"]      = _ => "DateTime.Now.TimeOfDay",
            ["Timer"]     = _ => "DateTime.Now.TimeOfDay.TotalSeconds",
            ["DateAdd"]   = a => $"/* DateAdd */ {RawArgs(a)} /* TODO: use DateTime.Add */",
            ["DateDiff"]  = a => $"/* DateDiff */ {RawArgs(a)} /* TODO: use TimeSpan */",
            ["DatePart"]  = a => $"/* DatePart */ {RawArgs(a)} /* TODO */",
            ["Year"]      = a => $"{One(a)}.Year",
            ["Month"]     = a => $"{One(a)}.Month",
            ["Day"]       = a => $"{One(a)}.Day",
            ["Hour"]      = a => $"{One(a)}.Hour",
            ["Minute"]    = a => $"{One(a)}.Minute",
            ["Second"]    = a => $"{One(a)}.Second",
            ["DateSerial"] = a => $"new DateTime({Arg(a,0)}, {Arg(a,1)}, {Arg(a,2)})",

            // Check functions
            ["IsNull"]    = a => $"({One(a)} == null)",
            ["IsEmpty"]   = a => $"({One(a)} == null)",
            ["IsNothing"] = a => $"({One(a)} == null)",
            ["IsNumeric"] = a => $"double.TryParse({One(a)}.ToString(), out _)",
            ["IsDate"]    = a => $"DateTime.TryParse({One(a)}.ToString(), out _)",
            ["IsArray"]   = a => $"({One(a)} is System.Array)",
            ["TypeName"]  = a => $"{One(a)}.GetType().Name",

            // Array
            ["UBound"]    = a => a.Length > 1
                                 ? $"{Arg(a,0)}.GetUpperBound({Arg(a,1)} - 1)"
                                 : $"({Arg(a,0)}.Length - 1)",
            ["LBound"]    = _ => "0",
            ["Array"]     = a => $"new object[] {{ {RawArgs(a)} }}",

            // Debug
            ["Debug.Print"] = a => $"System.Diagnostics.Debug.WriteLine({RawArgs(a)})",

            // Misc
            ["IIf"]       = a => $"({Arg(a,0)} ? {Arg(a,1)} : {Arg(a,2)})",
            ["Nz"]        = a => $"({Arg(a,0)} ?? {(a.Length > 1 ? Arg(a,1) : "\"\"")}) /* Nz */",
            ["Choose"]    = a => $"/* Choose */ {RawArgs(a)} /* TODO */",
            ["Switch"]    = a => $"/* Switch */ {RawArgs(a)} /* TODO */",
        };

    // VB6 constants → C# equivalents
    private static readonly Dictionary<string, string> Constants =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["vbCrLf"]          = "\"\\r\\n\"",
            ["vbCr"]            = "\"\\r\"",
            ["vbLf"]            = "\"\\n\"",
            ["vbTab"]           = "\"\\t\"",
            ["vbNullString"]    = "null",
            ["vbNullChar"]      = "'\\0'",
            ["vbTrue"]          = "true",
            ["vbFalse"]         = "false",
            ["vbObjectError"]   = "unchecked((int)0x80040000)",
            ["True"]            = "true",
            ["False"]           = "false",
            ["Nothing"]         = "null",
            ["Empty"]           = "null",
            ["Null"]            = "null",
        };

    public static bool TryGetFunction(string name, string[] args, out string result)
    {
        // Strip VB6 type-suffix '$' (e.g. Mid$ → Mid, Left$ → Left, Chr$ → Chr)
        string lookup = name.EndsWith('$') ? name[..^1] : name;
        if (Functions.TryGetValue(lookup, out var fn))
        {
            result = fn(args);
            return true;
        }
        result = "";
        return false;
    }

    public static bool TryGetConstant(string name, out string result) =>
        Constants.TryGetValue(name, out result!);

    private static string One(string[] a) => a.Length > 0 ? a[0] : "";
    private static string Arg(string[] a, int i) => i < a.Length ? a[i] : "";
    private static string RawArgs(string[] a) => string.Join(", ", a);
}
