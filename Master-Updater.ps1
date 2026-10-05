<#
.SYNOPSIS
    Excel Master Updater - Master file child records synchronization and update tool.

.DESCRIPTION
    Maintains one master ('base') Excel file of records. Incoming change files
    with arbitrary column layouts are mapped, fingerprint-matched, and reviewed
    row-by-row. Accepted new rows are appended, accepted changes are updated in-place,
    and all modifications are logged with exact cell references and stamped metadata.

.PARAMETER BaseFilePath
    Absolute or relative path to the master base Excel file (.xlsx) or CSV file (.csv).
    Pre-loads the base file into the application UI.

.PARAMETER IncomingPath
    Absolute or relative path to the incoming updates Excel file (.xlsx) or CSV file (.csv).
    Pre-loads the incoming file into the application UI.

.PARAMETER BaseSheet
    Name of the target worksheet in the base file. Defaults to the first available sheet.

.PARAMETER IncomingSheet
    Name of the source worksheet in the incoming file. Defaults to the first available sheet.

.PARAMETER ConfigPath
    Custom path to application configuration file. Defaults to '%APPDATA%\MasterUpdater\config.json'.

.PARAMETER NonInteractive
    Runs in headless mode without displaying the interactive WPF window.
    Used for automated pipeline verification and testing.

.PARAMETER Headless
    Runs the automated comparison, reporting, and optional write-back engine without opening the WPF interface.

.PARAMETER AutoAccept
    Batch acceptance policy when running in headless mode: 'None', 'AllNonAmbiguous', 'OnlyChanged', 'OnlyNew'.
    Defaults to 'None'.

.PARAMETER ExportReportPath
    Destination file path to export a discrepancy report (.html, .xlsx, or .csv).

.PARAMETER SummaryJsonPath
    Destination file path to export machine-readable execution summary JSON with record counts and write metrics.
    Defaults to '<LogDirectory>\last_execution_summary.json' if not explicitly passed.

.PARAMETER Quiet
    Suppresses console host output during headless execution.

.EXAMPLE
    pwsh -File .\Master-Updater.ps1
    Launches the application graphical interface.

.EXAMPLE
    pwsh -File .\Master-Updater.ps1 -BaseFilePath "C:\Data\Master.xlsx" -IncomingPath "C:\Data\Incoming.xlsx"
    Launches the application with files pre-loaded.

.EXAMPLE
    pwsh -File .\Master-Updater.ps1 -BaseFilePath "C:\Data\Master.xlsx" -IncomingPath "C:\Data\Incoming.csv" -Headless -AutoAccept AllNonAmbiguous -ExportReportPath "C:\Reports\diff.html" -SummaryJsonPath "C:\Logs\summary.json"
    Runs unattended batch synchronization with auto-accept, generates an HTML diff report, and exports summary JSON.

.OUTPUTS
    System.Management.Automation.PSCustomObject when running in -Headless mode; System.Void when running GUI.

.NOTES
    Compatible with Windows PowerShell 5.1 and PowerShell 7.x.
    Zero external dependencies (pure .NET OpenXML zip + XML).
    Compilable via ps2exe.

.LINK
    https://github.com/
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$BaseFilePath,

    [Parameter(Mandatory = $false)]
    [string]$IncomingPath,

    [Parameter(Mandatory = $false)]
    [string]$BaseSheet,

    [Parameter(Mandatory = $false)]
    [string]$IncomingSheet,

    [Parameter(Mandatory = $false)]
    [string]$ConfigPath,

    [Parameter(Mandatory = $false)]
    [switch]$NonInteractive,

    [Parameter(Mandatory = $false)]
    [switch]$Headless,

    [Parameter(Mandatory = $false)]
    [ValidateSet('None', 'AllNonAmbiguous', 'OnlyChanged', 'OnlyNew')]
    [string]$AutoAccept = 'None',

    [Parameter(Mandatory = $false)]
    [string]$ExportReportPath,

    [Parameter(Mandatory = $false)]
    [string]$SummaryJsonPath,

    [Parameter(Mandatory = $false)]
    [switch]$Quiet
)

try { Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction SilentlyContinue } catch { }

# ==============================================================================
# Region 1: Dependencies & Embedded C#
# ==============================================================================
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Xml
Add-Type -AssemblyName System.Xml.Linq

# Cached BrushConverter to avoid redundant allocations across UI renders
$script:BrushConverter = if ('System.Windows.Media.BrushConverter' -as [type]) { [System.Windows.Media.BrushConverter]::new() } else { $null }

# DwmHelper for dark mode window title bar
if (-not ('DwmHelper' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class DwmHelper {
    [DllImport("dwmapi.dll", PreserveSig = true)]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
}
"@ -ErrorAction SilentlyContinue
}

# ConsoleHelper for attaching to caller's console in headless/compiled mode
if (-not ('ConsoleHelper' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class ConsoleHelper {
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool AttachConsole(int dwProcessId);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool AllocConsole();

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool FreeConsole();
}
"@ -ErrorAction SilentlyContinue
}

function Set-WindowDwmTheme {
    param(
        [IntPtr]$Hwnd,
        [bool]$IsDark
    )
    try {
        if ($Hwnd -and $Hwnd -ne [IntPtr]::Zero) {
            $val = if ($IsDark) { 1 } else { 0 }
            $res = [DwmHelper]::DwmSetWindowAttribute($Hwnd, 20, [ref]$val, 4)
            if ($res -ne 0) {
                [DwmHelper]::DwmSetWindowAttribute($Hwnd, 19, [ref]$val, 4)
            }
        }
    } catch { }
}

# FastDiffHelper for value normalization, equality check, and merging
if (-not ('FastDiffHelper' -as [type])) {
    $diffCSharp = @"
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

public static class FastDiffHelper {
    private static readonly Regex _rgxSpecial = new Regex(@"[^\p{L}\p{Nd}\s]", RegexOptions.Compiled);
    private static readonly Regex _rgxWhitespace = new Regex(@"\s+", RegexOptions.Compiled);

    public static string Normalize(string val, bool ignoreCase, bool trim, bool ignoreSpecial, bool ignoreAllWhitespace) {
        if (string.IsNullOrEmpty(val)) return string.Empty;
        string v = val;
        if (trim) v = v.Trim();
        if (ignoreCase) v = v.ToLowerInvariant();
        if (ignoreSpecial) v = _rgxSpecial.Replace(v, string.Empty);
        if (ignoreAllWhitespace) {
            v = _rgxWhitespace.Replace(v, string.Empty);
        } else {
            v = _rgxWhitespace.Replace(v, " ").Trim();
        }
        return v;
    }

    public static bool AreEqual(string bRaw, string uRaw, bool ignoreCase, bool trim, bool ignoreSpecial, bool ignoreAllSpaces) {
        if (string.Equals(bRaw, uRaw, StringComparison.Ordinal)) return true;
        if (bRaw == null) bRaw = string.Empty;
        if (uRaw == null) uRaw = string.Empty;
        if (bRaw.Length == 0 && uRaw.Length == 0) return true;

        if (ignoreCase && !ignoreSpecial && !ignoreAllSpaces && !trim) {
            return string.Equals(bRaw, uRaw, StringComparison.OrdinalIgnoreCase);
        }
        if (ignoreCase && !ignoreSpecial && !ignoreAllSpaces && trim) {
            return string.Equals(bRaw.Trim(), uRaw.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        string bNorm = Normalize(bRaw, ignoreCase, trim, ignoreSpecial, ignoreAllSpaces);
        string uNorm = Normalize(uRaw, ignoreCase, trim, ignoreSpecial, ignoreAllSpaces);
        return string.Equals(bNorm, uNorm, StringComparison.Ordinal);
    }

    public static string MergeValues(IList<string> values, string mergeMode, string separator, bool trim) {
        if (values == null || values.Count == 0) return string.Empty;
        if (string.Equals(mergeMode, "Exact", StringComparison.OrdinalIgnoreCase)) {
            string v = values[0] ?? string.Empty;
            return trim ? v.Trim() : v;
        }
        if (string.Equals(mergeMode, "FirstNonEmpty", StringComparison.OrdinalIgnoreCase)) {
            for (int i = 0; i < values.Count; i++) {
                string v = values[i];
                if (!string.IsNullOrEmpty(v)) {
                    if (trim) v = v.Trim();
                    if (!string.IsNullOrEmpty(v)) return v;
                }
            }
            return string.Empty;
        }
        StringBuilder sb = new StringBuilder();
        string sep = separator ?? string.Empty;
        for (int i = 0; i < values.Count; i++) {
            string v = values[i] ?? string.Empty;
            if (trim) v = v.Trim();
            if (i > 0) sb.Append(sep);
            sb.Append(v);
        }
        return sb.ToString();
    }
}
"@
    Add-Type -TypeDefinition $diffCSharp -Language CSharp
}

# FastExcelHelper for fast OpenXML reading and exporting
if (-not ('FastExcelHelper' -as [type])) {
    $excelCSharp = @"
using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Globalization;
using System.Management.Automation;

public static class FastExcelHelper {
    public static int CellRefToColIndex(string cellRef, int fallback) {
        if (string.IsNullOrEmpty(cellRef)) return fallback;
        int col = 0;
        int i = 0;
        while (i < cellRef.Length && char.IsLetter(cellRef[i])) {
            col = col * 26 + (char.ToUpperInvariant(cellRef[i]) - 'A' + 1);
            i++;
        }
        return (col > 0) ? (col - 1) : fallback;
    }

    public static string ColIndexToName(int index) {
        int dividend = index + 1;
        string colName = "";
        while (dividend > 0) {
            int modulo = (dividend - 1) % 26;
            colName = Convert.ToChar(65 + modulo) + colName;
            dividend = (dividend - modulo) / 26;
        }
        return colName;
    }

    public static string EscapeXml(string val) {
        if (string.IsNullOrEmpty(val)) return string.Empty;
        return val.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;");
    }

    private static string ResolveSheetTarget(ZipArchive zip, string sheetName) {
        var wbEntry = zip.GetEntry("xl/workbook.xml");
        if (wbEntry == null) return "xl/worksheets/sheet1.xml";
        string rId = null;
        using (var s = wbEntry.Open())
        using (var xr = XmlReader.Create(s)) {
            while (xr.Read()) {
                if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "sheet") {
                    string name = xr.GetAttribute("name");
                    string id = xr.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
                             ?? xr.GetAttribute("r:id");
                    if (string.Equals(name, sheetName, StringComparison.OrdinalIgnoreCase)) {
                        rId = id;
                        break;
                    }
                    if (rId == null) rId = id;
                }
            }
        }
        if (string.IsNullOrEmpty(rId)) return "xl/worksheets/sheet1.xml";

        var relsEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");
        if (relsEntry == null) return "xl/worksheets/sheet1.xml";
        using (var s = relsEntry.Open())
        using (var xr = XmlReader.Create(s)) {
            while (xr.Read()) {
                if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "Relationship") {
                    if (xr.GetAttribute("Id") == rId) {
                        string target = xr.GetAttribute("Target");
                        if (!target.StartsWith("xl/")) target = "xl/" + target.TrimStart('/');
                        return target;
                    }
                }
            }
        }
        return "xl/worksheets/sheet1.xml";
    }

    /// <summary>
    /// Detects character encoding of a file by inspecting BOM headers, UTF-8 byte sequences,
    /// and falling back to Windows-1250 (Central European ANSI) if high bytes are non-UTF-8.
    /// </summary>
    public static Encoding DetectFileEncoding(string filePath) {
        if (!File.Exists(filePath)) return Encoding.UTF8;
        byte[] buffer = new byte[8192];
        int bytesRead = 0;
        using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
            bytesRead = fs.Read(buffer, 0, buffer.Length);
        }
        if (bytesRead >= 3 && buffer[0] == 0xEF && buffer[1] == 0xBB && buffer[2] == 0xBF) return Encoding.UTF8;
        if (bytesRead >= 2 && buffer[0] == 0xFF && buffer[1] == 0xFE) return Encoding.Unicode;
        if (bytesRead >= 2 && buffer[0] == 0xFE && buffer[1] == 0xFF) return Encoding.BigEndianUnicode;

        // Check if buffer contains valid UTF-8 sequences
        bool isValidUtf8 = true;
        bool hasHighBytes = false;
        for (int i = 0; i < bytesRead; i++) {
            byte b = buffer[i];
            if (b > 0x7F) {
                hasHighBytes = true;
                if ((b & 0xE0) == 0xC0) {
                    if (i + 1 >= bytesRead || (buffer[i + 1] & 0xC0) != 0x80) { isValidUtf8 = false; break; }
                    i += 1;
                } else if ((b & 0xF0) == 0xE0) {
                    if (i + 2 >= bytesRead || (buffer[i + 1] & 0xC0) != 0x80 || (buffer[i + 2] & 0xC0) != 0x80) { isValidUtf8 = false; break; }
                    i += 2;
                } else if ((b & 0xF8) == 0xF0) {
                    if (i + 3 >= bytesRead || (buffer[i + 1] & 0xC0) != 0x80 || (buffer[i + 2] & 0xC0) != 0x80 || (buffer[i + 3] & 0xC0) != 0x80) { isValidUtf8 = false; break; }
                    i += 3;
                } else {
                    isValidUtf8 = false;
                    break;
                }
            }
        }
        if (!isValidUtf8 && hasHighBytes) {
            try {
                return Encoding.GetEncoding(1250);
            } catch {
                return Encoding.Default;
            }
        }
        return Encoding.UTF8;
    }

    /// <summary>
    /// Analyzes the first line of a CSV file to automatically identify delimiter: ';', ',', '\t', or '|',
    /// ignoring delimiters enclosed in double quotes. Defaults to ';' if ambiguous.
    /// </summary>
    public static char DetectCsvDelimiter(string firstLine) {
        if (string.IsNullOrEmpty(firstLine)) return ';';
        int semicolons = 0, commas = 0, tabs = 0, pipes = 0;
        bool inQuotes = false;
        for (int i = 0; i < firstLine.Length; i++) {
            char c = firstLine[i];
            if (c == '"') inQuotes = !inQuotes;
            else if (!inQuotes) {
                if (c == ';') semicolons++;
                else if (c == ',') commas++;
                else if (c == '\t') tabs++;
                else if (c == '|') pipes++;
            }
        }
        if (semicolons >= commas && semicolons >= tabs && semicolons >= pipes && semicolons > 0) return ';';
        if (commas >= semicolons && commas >= tabs && commas >= pipes && commas > 0) return ',';
        if (tabs >= semicolons && tabs >= commas && tabs >= pipes && tabs > 0) return '\t';
        if (pipes > 0) return '|';
        return ';';
    }

    public static List<string> ParseCsvLine(string line, char delimiter) {
        var fields = new List<string>();
        if (line == null) return fields;
        var sb = new StringBuilder();
        bool inQuotes = false;
        for (int i = 0; i < line.Length; i++) {
            char c = line[i];
            if (c == '"') {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"') {
                    sb.Append('"');
                    i++;
                } else {
                    inQuotes = !inQuotes;
                }
            } else if (c == delimiter && !inQuotes) {
                fields.Add(sb.ToString().Trim());
                sb.Length = 0;
            } else {
                sb.Append(c);
            }
        }
        fields.Add(sb.ToString().Trim());
        return fields;
    }

    public static List<string> GetSheetNames(string filePath) {
        var list = new List<string>();
        if (!File.Exists(filePath)) return list;
        if (filePath.EndsWith(".csv", StringComparison.OrdinalIgnoreCase)) {
            list.Add("CSV");
            return list;
        }
        using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        using (var zip = new ZipArchive(fs, ZipArchiveMode.Read)) {
            var wb = zip.GetEntry("xl/workbook.xml");
            if (wb == null) return list;
            using (var s = wb.Open())
            using (var xr = XmlReader.Create(s)) {
                while (xr.Read()) {
                    if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "sheet") {
                        string name = xr.GetAttribute("name");
                        if (!string.IsNullOrEmpty(name)) list.Add(name);
                    }
                }
            }
        }
        return list;
    }

    private static List<string> GetSharedStrings(ZipArchive zip) {
        var sst = new List<string>();
        var entry = zip.GetEntry("xl/sharedStrings.xml");
        if (entry == null) return sst;
        using (var s = entry.Open())
        using (var xr = XmlReader.Create(s)) {
            var currentStr = new StringBuilder();
            bool inT = false;
            while (xr.Read()) {
                if (xr.NodeType == XmlNodeType.Element) {
                    if (xr.LocalName == "si") {
                        currentStr.Length = 0;
                    } else if (xr.LocalName == "t") {
                        inT = true;
                    }
                } else if (inT && (xr.NodeType == XmlNodeType.Text || xr.NodeType == XmlNodeType.SignificantWhitespace)) {
                    currentStr.Append(xr.Value);
                } else if (xr.NodeType == XmlNodeType.EndElement) {
                    if (xr.LocalName == "t") {
                        inT = false;
                    } else if (xr.LocalName == "si") {
                        sst.Add(currentStr.ToString());
                    }
                }
            }
        }
        return sst;
    }

    public static List<string> GetHeaders(string filePath, string sheetName) {
        var headers = new List<string>();
        if (!File.Exists(filePath)) return headers;
        if (filePath.EndsWith(".csv", StringComparison.OrdinalIgnoreCase)) {
            Encoding enc = DetectFileEncoding(filePath);
            using (var sr = new StreamReader(filePath, enc, true)) {
                string firstLine = null;
                while (!sr.EndOfStream) {
                    string line = sr.ReadLine();
                    if (!string.IsNullOrWhiteSpace(line)) { firstLine = line; break; }
                }
                if (firstLine != null) {
                    char d = DetectCsvDelimiter(firstLine);
                    return ParseCsvLine(firstLine, d);
                }
            }
            return headers;
        }
        using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        using (var zip = new ZipArchive(fs, ZipArchiveMode.Read)) {
            var sst = GetSharedStrings(zip);
            string target = ResolveSheetTarget(zip, sheetName);
            var entry = zip.GetEntry(target);
            if (entry == null) return headers;
            using (var s = entry.Open())
            using (var xr = XmlReader.Create(s)) {
                bool inRow1 = false;
                bool inCell = false;
                bool inVal = false;
                string cellType = null;
                var sb = new StringBuilder();
                int col = -1;
                var dict = new SortedDictionary<int, string>();

                while (xr.Read()) {
                    if (xr.NodeType == XmlNodeType.Element) {
                        if (xr.LocalName == "row") {
                            string r = xr.GetAttribute("r");
                            if (r == "1" || inRow1 == false) inRow1 = true;
                            else if (inRow1) break;
                        } else if (inRow1 && xr.LocalName == "c") {
                            inCell = true;
                            string r = xr.GetAttribute("r");
                            cellType = xr.GetAttribute("t");
                            col = CellRefToColIndex(r, col + 1);
                            sb.Length = 0;
                        } else if (inCell && (xr.LocalName == "v" || xr.LocalName == "t")) {
                            inVal = true;
                        }
                    } else if (inVal && (xr.NodeType == XmlNodeType.Text || xr.NodeType == XmlNodeType.SignificantWhitespace)) {
                        sb.Append(xr.Value);
                    } else if (xr.NodeType == XmlNodeType.EndElement) {
                        if (xr.LocalName == "v" || xr.LocalName == "t") {
                            inVal = false;
                        } else if (xr.LocalName == "c") {
                            inCell = false;
                            string val = sb.ToString();
                            if (cellType == "s") {
                                int sstIdx;
                                if (int.TryParse(val, out sstIdx) && sstIdx >= 0 && sstIdx < sst.Count) {
                                    val = sst[sstIdx];
                                }
                            }
                            dict[col] = val.Trim();
                        } else if (xr.LocalName == "row" && inRow1) {
                            break;
                        }
                    }
                }
                foreach (var v in dict.Values) headers.Add(v);
            }
        }
        return headers;
    }

    public static List<PSObject> ReadSheet(string filePath, string sheetName, int endRow = 0) {
        var results = new List<PSObject>();
        if (!File.Exists(filePath)) return results;
        if (filePath.EndsWith(".csv", StringComparison.OrdinalIgnoreCase)) {
            Encoding enc = DetectFileEncoding(filePath);
            using (var sr = new StreamReader(filePath, enc, true)) {
                string firstLine = null;
                int physicalRow = 0;
                while (!sr.EndOfStream) {
                    string line = sr.ReadLine();
                    physicalRow++;
                    if (!string.IsNullOrWhiteSpace(line)) { firstLine = line; break; }
                }
                if (firstLine == null) return results;
                char d = DetectCsvDelimiter(firstLine);
                var headers = ParseCsvLine(firstLine, d);
                for (int i = 0; i < headers.Count; i++) {
                    if (string.IsNullOrEmpty(headers[i])) headers[i] = "Column" + (i + 1);
                }

                while (!sr.EndOfStream) {
                    string line = sr.ReadLine();
                    physicalRow++;
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    var fields = ParseCsvLine(line, d);
                    var pso = new PSObject();
                    pso.Properties.Add(new PSNoteProperty("_RowNumber", physicalRow));
                    for (int h = 0; h < headers.Count; h++) {
                        string val = (h < fields.Count) ? fields[h] : string.Empty;
                        pso.Properties.Add(new PSNoteProperty(headers[h], val));
                    }
                    results.Add(pso);
                    if (endRow > 0 && results.Count >= endRow) break;
                }
            }
            return results;
        }
        using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        using (var zip = new ZipArchive(fs, ZipArchiveMode.Read)) {
            var sst = GetSharedStrings(zip);
            string target = ResolveSheetTarget(zip, sheetName);
            var entry = zip.GetEntry(target);
            if (entry == null) return results;

            using (var s = entry.Open())
            using (var xr = XmlReader.Create(s)) {
                var headers = new List<string>();
                var headerCols = new List<int>();
                bool foundHeaders = false;
                int currentPhysicalRow = 0;
                int estimatedRow = 0;
                var rowCells = new Dictionary<int, string>();
                int currentCellCol = -1;
                string currentCellType = null;
                var currentVal = new StringBuilder();
                bool inRow = false;
                bool inCell = false;
                bool inVal = false;

                while (xr.Read()) {
                    if (xr.NodeType == XmlNodeType.Element) {
                        if (xr.LocalName == "row") {
                            inRow = true;
                            rowCells.Clear();
                            estimatedRow++;
                            string rAttr = xr.GetAttribute("r");
                            if (!int.TryParse(rAttr, out currentPhysicalRow) || currentPhysicalRow <= 0) {
                                currentPhysicalRow = estimatedRow;
                            } else {
                                estimatedRow = currentPhysicalRow;
                            }
                        } else if (inRow && xr.LocalName == "c") {
                            inCell = true;
                            string r = xr.GetAttribute("r");
                            currentCellType = xr.GetAttribute("t");
                            currentCellCol = CellRefToColIndex(r, currentCellCol + 1);
                            currentVal.Length = 0;
                        } else if (inCell && (xr.LocalName == "v" || xr.LocalName == "t")) {
                            inVal = true;
                        }
                    } else if (inVal && (xr.NodeType == XmlNodeType.Text || xr.NodeType == XmlNodeType.SignificantWhitespace)) {
                        currentVal.Append(xr.Value);
                    } else if (xr.NodeType == XmlNodeType.EndElement) {
                        if (xr.LocalName == "v" || xr.LocalName == "t") {
                            inVal = false;
                        } else if (xr.LocalName == "c") {
                            inCell = false;
                            if (currentCellCol >= 0) {
                                string val = currentVal.ToString();
                                if (currentCellType == "s") {
                                    int sstIdx;
                                    if (int.TryParse(val, out sstIdx) && sstIdx >= 0 && sstIdx < sst.Count) {
                                        val = sst[sstIdx];
                                    }
                                } else if (currentCellType == "b") {
                                    val = (val == "1") ? "True" : "False";
                                }
                                rowCells[currentCellCol] = val;
                            }
                        } else if (xr.LocalName == "row") {
                            inRow = false;
                            if (!foundHeaders) {
                                int maxCol = -1;
                                foreach (var k in rowCells.Keys) if (k > maxCol) maxCol = k;
                                for (int c = 0; c <= maxCol; c++) {
                                    string h = rowCells.ContainsKey(c) ? rowCells[c].Trim() : ("Column" + (c + 1));
                                    if (string.IsNullOrEmpty(h)) h = "Column" + (c + 1);
                                    headers.Add(h);
                                    headerCols.Add(c);
                                }
                                foundHeaders = true;
                            } else {
                                var pso = new PSObject();
                                pso.Properties.Add(new PSNoteProperty("_RowNumber", currentPhysicalRow));
                                for (int h = 0; h < headers.Count; h++) {
                                    string hName = headers[h];
                                    int cIdx = headerCols[h];
                                    string val = rowCells.ContainsKey(cIdx) ? rowCells[cIdx] : string.Empty;
                                    pso.Properties.Add(new PSNoteProperty(hName, val));
                                }
                                results.Add(pso);
                                if (endRow > 0 && results.Count >= endRow) break;
                            }
                        }
                    }
                }
            }
        }
        return results;
    }

    public static void ExportToExcel(string filePath, IEnumerable<object> rows, bool autoSize = true, bool freezeTopRow = true, bool boldTopRow = true) {
        string tempZip = Path.GetTempFileName();
        if (File.Exists(tempZip)) File.Delete(tempZip);

        var list = new List<PSObject>();
        foreach (var r in rows) {
            PSObject pso = r as PSObject;
            if (pso != null) list.Add(pso);
            else if (r != null) list.Add(PSObject.AsPSObject(r));
        }

        var headers = new List<string>();
        if (list.Count > 0) {
            foreach (var prop in list[0].Properties) {
                if (!prop.Name.StartsWith("_")) headers.Add(prop.Name);
            }
        }

        using (var fs = new FileStream(tempZip, FileMode.Create, FileAccess.Write))
        using (var zip = new ZipArchive(fs, ZipArchiveMode.Create)) {
            var ct = zip.CreateEntry("[Content_Types].xml");
            using (var sw = new StreamWriter(ct.Open(), Encoding.UTF8)) {
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/></Types>");
            }
            var rels = zip.CreateEntry("_rels/.rels");
            using (var sw = new StreamWriter(rels.Open(), Encoding.UTF8)) {
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
            }
            var wbRels = zip.CreateEntry("xl/_rels/workbook.xml.rels");
            using (var sw = new StreamWriter(wbRels.Open(), Encoding.UTF8)) {
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships>");
            }
            var wb = zip.CreateEntry("xl/workbook.xml");
            using (var sw = new StreamWriter(wb.Open(), Encoding.UTF8)) {
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
            }
            var st = zip.CreateEntry("xl/styles.xml");
            using (var sw = new StreamWriter(st.Open(), Encoding.UTF8)) {
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts count=\"2\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font><font><b/><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"2\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/></cellXfs></styleSheet>");
            }
            var ws = zip.CreateEntry("xl/worksheets/sheet1.xml");
            using (var sw = new StreamWriter(ws.Open(), Encoding.UTF8)) {
                int totalRows = list.Count + 1;
                string lastCol = ColIndexToName(Math.Max(0, headers.Count - 1));
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><dimension ref=\"A1:" + lastCol + totalRows + "\"/>");
                if (freezeTopRow) {
                    sw.Write("<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/></sheetView></sheetViews>");
                }
                sw.Write("<sheetData><row r=\"1\">");
                for (int c = 0; c < headers.Count; c++) {
                    string refId = ColIndexToName(c) + "1";
                    string sAttr = boldTopRow ? " s=\"1\"" : "";
                    sw.Write("<c r=\"" + refId + "\"" + sAttr + " t=\"inlineStr\"><is><t>" + EscapeXml(headers[c]) + "</t></is></c>");
                }
                sw.Write("</row>");
                for (int r = 0; r < list.Count; r++) {
                    int rNum = r + 2;
                    sw.Write("<row r=\"" + rNum + "\">");
                    var pso = list[r];
                    for (int c = 0; c < headers.Count; c++) {
                        string h = headers[c];
                        var pVal = pso.Properties[h] != null ? pso.Properties[h].Value : null;
                        string vStr = pVal != null ? pVal.ToString() : "";
                        string refId = ColIndexToName(c) + rNum;
                        sw.Write("<c r=\"" + refId + "\" t=\"inlineStr\"><is><t>" + EscapeXml(vStr) + "</t></is></c>");
                    }
                    sw.Write("</row>");
                }
                sw.Write("</sheetData></worksheet>");
            }
        }
        string dir = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);
        File.Copy(tempZip, filePath, true);
        try { File.Delete(tempZip); } catch { }
    }
}
"@
    if ($PSVersionTable.PSVersion.Major -le 5) {
        Add-Type -TypeDefinition $excelCSharp -ReferencedAssemblies @(
            'System.Xml',
            'System.IO.Compression',
            'System.IO.Compression.FileSystem',
            'System.Core',
            ([PSObject].Assembly.Location)
        ) -Language CSharp
    } else {
        Add-Type -TypeDefinition $excelCSharp -Language CSharp
    }
}

# EditExcelHelper for in-place XML cell patching, row appending, and structure guards
if (-not ('RowOp' -as [type])) {
    $editCSharp = @"
using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Linq;

public class RowOp {
    public string Type; // "PatchCell" or "AppendRow"
    public int RowNumber;
    public Dictionary<int, string> Cells = new Dictionary<int, string>();
}

public static class EditExcelHelper {
    private static readonly XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    public static int CellRefToColIndex(string cellRef, int fallback) {
        if (string.IsNullOrEmpty(cellRef)) return fallback;
        int col = 0; int i = 0;
        while (i < cellRef.Length && char.IsLetter(cellRef[i])) {
            col = col * 26 + (char.ToUpperInvariant(cellRef[i]) - 'A' + 1);
            i++;
        }
        return (col > 0) ? (col - 1) : fallback;
    }

    public static string ColIndexToName(int index) {
        int dividend = index + 1;
        string colName = "";
        while (dividend > 0) {
            int modulo = (dividend - 1) % 26;
            colName = Convert.ToChar(65 + modulo) + colName;
            dividend = (dividend - modulo) / 26;
        }
        return colName;
    }

    public static void WriteChanges(string filePath, string sheetName, IList<RowOp> ops) {
        if (!File.Exists(filePath)) throw new FileNotFoundException("Base file not found", filePath);
        string tempPath = filePath + ".tmp.xlsx";
        if (File.Exists(tempPath)) File.Delete(tempPath);

        string targetSheetPath = null;
        XDocument doc = null;

        try {
            using (var inStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var zipIn = new ZipArchive(inStream, ZipArchiveMode.Read)) {
                // Locate sheet
                var wbEntry = zipIn.GetEntry("xl/workbook.xml");
                string rId = null;
                if (wbEntry != null) {
                    using (var s = wbEntry.Open())
                    using (var xr = XmlReader.Create(s)) {
                        while (xr.Read()) {
                            if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "sheet") {
                                string name = xr.GetAttribute("name");
                                string id = xr.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
                                         ?? xr.GetAttribute("r:id");
                                if (string.Equals(name, sheetName, StringComparison.OrdinalIgnoreCase)) {
                                    rId = id; break;
                                }
                                if (rId == null) rId = id;
                            }
                        }
                    }
                }
                var relsEntry = zipIn.GetEntry("xl/_rels/workbook.xml.rels");
                if (relsEntry != null && !string.IsNullOrEmpty(rId)) {
                    using (var s = relsEntry.Open())
                    using (var xr = XmlReader.Create(s)) {
                        while (xr.Read()) {
                            if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "Relationship") {
                                if (xr.GetAttribute("Id") == rId) {
                                    string target = xr.GetAttribute("Target");
                                    if (!target.StartsWith("xl/")) target = "xl/" + target.TrimStart('/');
                                    targetSheetPath = target;
                                    break;
                                }
                            }
                        }
                    }
                }
                if (string.IsNullOrEmpty(targetSheetPath)) targetSheetPath = "xl/worksheets/sheet1.xml";

                var sheetEntry = zipIn.GetEntry(targetSheetPath);
                if (sheetEntry == null) throw new InvalidOperationException("Target worksheet not found: " + targetSheetPath);

                using (var sheetStream = sheetEntry.Open()) {
                    doc = XDocument.Load(sheetStream);
                }

                // GUARD: Scan for tableParts, pivotTable, and shared formulas without full ToString() memory overhead
                // P5: Returns a warning string instead of throwing so callers can fall back to SafeRewrite
                bool hasUnsupported = doc.Descendants().Any(e =>
                    e.Name.LocalName.Equals("tableParts", StringComparison.OrdinalIgnoreCase) ||
                    e.Name.LocalName.Equals("tablePart", StringComparison.OrdinalIgnoreCase) ||
                    e.Name.LocalName.Equals("pivotTable", StringComparison.OrdinalIgnoreCase) ||
                    (e.Name.LocalName == "f" && string.Equals((string)e.Attribute("t"), "shared", StringComparison.OrdinalIgnoreCase))
                );
                if (hasUnsupported) {
                    throw new InvalidOperationException("INPLACE_GUARD_WARN: Unsupported structure detected (tableParts, pivotTable, or shared formulas). Falling back to SafeRewrite mode.");
                }

                var sheetData = doc.Root.Element(ns + "sheetData");
                if (sheetData == null) throw new InvalidOperationException("sheetData element not found in worksheet");

                foreach (var op in ops) {
                    if (op.Type == "PatchCell") {
                        var rowElem = sheetData.Elements(ns + "row").FirstOrDefault(r => (int?)r.Attribute("r") == op.RowNumber);
                        if (rowElem == null) {
                            rowElem = new XElement(ns + "row", new XAttribute("r", op.RowNumber));
                            var nextRow = sheetData.Elements(ns + "row").FirstOrDefault(r => (int?)r.Attribute("r") > op.RowNumber);
                            if (nextRow != null) nextRow.AddBeforeSelf(rowElem);
                            else sheetData.Add(rowElem);
                        }

                        foreach (var kvp in op.Cells) {
                            int colIdx = kvp.Key;
                            string colLetter = ColIndexToName(colIdx);
                            string cellRef = colLetter + op.RowNumber;
                            string cellVal = kvp.Value ?? "";

                            var cElem = rowElem.Elements(ns + "c").FirstOrDefault(c => (string)c.Attribute("r") == cellRef);
                            if (cElem != null) {
                                string styleAttr = (string)cElem.Attribute("s");
                                cElem.RemoveNodes();
                                cElem.SetAttributeValue("t", "inlineStr");
                                if (!string.IsNullOrEmpty(styleAttr)) cElem.SetAttributeValue("s", styleAttr);

                                var isElem = new XElement(ns + "is", new XElement(ns + "t", cellVal));
                                if (cellVal.Length > 0 && (char.IsWhiteSpace(cellVal[0]) || char.IsWhiteSpace(cellVal[cellVal.Length - 1]))) {
                                    isElem.Element(ns + "t").SetAttributeValue(XNamespace.Xml + "space", "preserve");
                                }
                                cElem.Add(isElem);
                            } else {
                                var newC = new XElement(ns + "c",
                                    new XAttribute("r", cellRef),
                                    new XAttribute("t", "inlineStr"),
                                    new XElement(ns + "is", new XElement(ns + "t", cellVal))
                                );
                                if (cellVal.Length > 0 && (char.IsWhiteSpace(cellVal[0]) || char.IsWhiteSpace(cellVal[cellVal.Length - 1]))) {
                                    newC.Element(ns + "is").Element(ns + "t").SetAttributeValue(XNamespace.Xml + "space", "preserve");
                                }

                                var nextCell = rowElem.Elements(ns + "c").FirstOrDefault(c => {
                                    int cCol = CellRefToColIndex((string)c.Attribute("r"), -1);
                                    return cCol > colIdx;
                                });
                                if (nextCell != null) nextCell.AddBeforeSelf(newC);
                                else rowElem.Add(newC);
                            }
                        }
                    } else if (op.Type == "AppendRow") {
                        var rowElem = new XElement(ns + "row", new XAttribute("r", op.RowNumber));
                        var sortedCells = op.Cells.OrderBy(k => k.Key);
                        foreach (var kvp in sortedCells) {
                            int colIdx = kvp.Key;
                            string colLetter = ColIndexToName(colIdx);
                            string cellRef = colLetter + op.RowNumber;
                            string cellVal = kvp.Value ?? "";

                            var newC = new XElement(ns + "c",
                                new XAttribute("r", cellRef),
                                new XAttribute("t", "inlineStr"),
                                new XElement(ns + "is", new XElement(ns + "t", cellVal))
                            );
                            if (cellVal.Length > 0 && (char.IsWhiteSpace(cellVal[0]) || char.IsWhiteSpace(cellVal[cellVal.Length - 1]))) {
                                newC.Element(ns + "is").Element(ns + "t").SetAttributeValue(XNamespace.Xml + "space", "preserve");
                            }
                            rowElem.Add(newC);
                        }
                        sheetData.Add(rowElem);
                    }
                }

                // Update dimension using fast incremental bounds without scanning all cells
                var dimElem = doc.Root.Element(ns + "dimension");
                if (dimElem != null) {
                    int maxRow = 1;
                    int maxCol = 0;
                    string existingRef = (string)dimElem.Attribute("ref");
                    if (!string.IsNullOrEmpty(existingRef)) {
                        int colonIdx = existingRef.IndexOf(':');
                        string endRef = (colonIdx >= 0) ? existingRef.Substring(colonIdx + 1) : existingRef;
                        maxCol = CellRefToColIndex(endRef, 0);
                        int rStart = 0;
                        while (rStart < endRef.Length && char.IsLetter(endRef[rStart])) rStart++;
                        int parsedR;
                        if (int.TryParse(endRef.Substring(rStart), out parsedR)) maxRow = parsedR;
                    }
                    foreach (var op in ops) {
                        if (op.RowNumber > maxRow) maxRow = op.RowNumber;
                        foreach (var kvp in op.Cells) {
                            if (kvp.Key > maxCol) maxCol = kvp.Key;
                        }
                    }
                    string dimRef = "A1:" + ColIndexToName(maxCol) + maxRow;
                    dimElem.SetAttributeValue("ref", dimRef);
                }

                // Stream copy to temp archive
                using (var outStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None))
                using (var zipOut = new ZipArchive(outStream, ZipArchiveMode.Create)) {
                    foreach (var entry in zipIn.Entries) {
                        if (string.Equals(entry.FullName, targetSheetPath, StringComparison.OrdinalIgnoreCase)) {
                            var newSheetEntry = zipOut.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                            using (var nsStream = newSheetEntry.Open())
                            using (var xw = XmlWriter.Create(nsStream, new XmlWriterSettings {
                                Encoding = new UTF8Encoding(false),
                                OmitXmlDeclaration = false,
                                Indent = false
                            })) {
                                doc.Save(xw);
                            }
                        } else {
                            var copyEntry = zipOut.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                            using (var src = entry.Open())
                            using (var dst = copyEntry.Open()) {
                                src.CopyTo(dst);
                            }
                        }
                    }
                }
            }
        } catch (IOException ioEx) {
            throw new IOException("Base file is locked by another process (e.g. Microsoft Excel). Close the file and retry.\n(" + ioEx.Message + ")", ioEx);
        }

        // Verify readable before swap
        try {
            using (var chkStream = new FileStream(tempPath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var chkZip = new ZipArchive(chkStream, ZipArchiveMode.Read)) {
                var entry = chkZip.GetEntry(targetSheetPath);
                if (entry == null) throw new InvalidOperationException("Modified sheet missing from archive");
                using (var s = entry.Open()) {
                    var chkDoc = XDocument.Load(s);
                }
            }
        } catch (Exception ex) {
            if (File.Exists(tempPath)) File.Delete(tempPath);
            throw new InvalidOperationException("Verification of modified worksheet failed: " + ex.Message, ex);
        }

        // Atomic swap using transactional File.Replace on NTFS with fallback to Copy/Delete
        try {
            File.Replace(tempPath, filePath, null);
        } catch {
            File.Copy(tempPath, filePath, true);
            try { File.Delete(tempPath); } catch { }
        }
    }

    public static void AppendWorksheetLog(string filePath, string sheetName, string[] headers, System.Collections.IEnumerable rows) {
        if (!File.Exists(filePath)) throw new FileNotFoundException("Base file not found", filePath);
        if (rows == null) return;

        var rowList = new List<List<string>>();
        foreach (var rObj in rows) {
            if (rObj == null) continue;
            var vals = new List<string>();
            var ie = rObj as System.Collections.IEnumerable;
            if (ie != null && !(rObj is string)) {
                foreach (var cell in ie) {
                    vals.Add(cell != null ? cell.ToString() : "");
                }
            } else {
                vals.Add(rObj.ToString());
            }
            if (vals.Count > 0) rowList.Add(vals);
        }
        if (rowList.Count == 0) return;

        string tempPath = filePath + ".tmp.xlsx";
        if (File.Exists(tempPath)) File.Delete(tempPath);

        bool sheetExists = false;
        string targetSheetPath = null;
        string targetRelsPath = null;
        int newSheetId = 1;
        string newRId = "rId1";

        XDocument ctDoc = null;
        XDocument wbDoc = null;
        XDocument relsDoc = null;
        XDocument sheetDoc = null;

        var relNs = XNamespace.Get("http://schemas.openxmlformats.org/package/2006/relationships");
        var ctNs = XNamespace.Get("http://schemas.openxmlformats.org/package/2006/content-types");
        var offRelNs = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        try {
            using (var inStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var zipIn = new ZipArchive(inStream, ZipArchiveMode.Read)) {
                var wbEntry = zipIn.GetEntry("xl/workbook.xml");
                if (wbEntry == null) throw new InvalidOperationException("xl/workbook.xml not found in archive");
                using (var s = wbEntry.Open()) wbDoc = XDocument.Load(s);

                var relsEntry = zipIn.GetEntry("xl/_rels/workbook.xml.rels");
                if (relsEntry == null) throw new InvalidOperationException("xl/_rels/workbook.xml.rels not found in archive");
                using (var s = relsEntry.Open()) relsDoc = XDocument.Load(s);

                var ctEntry = zipIn.GetEntry("[Content_Types].xml");
                if (ctEntry == null) throw new InvalidOperationException("[Content_Types].xml not found in archive");
                using (var s = ctEntry.Open()) ctDoc = XDocument.Load(s);

                var sheetsElem = wbDoc.Root.Element(ns + "sheets");
                if (sheetsElem == null) {
                    sheetsElem = new XElement(ns + "sheets");
                    wbDoc.Root.Add(sheetsElem);
                }

                string existingRId = null;
                int maxSheetId = 0;
                foreach (var s in sheetsElem.Elements(ns + "sheet")) {
                    int sId = (int?)s.Attribute("sheetId") ?? 0;
                    if (sId > maxSheetId) maxSheetId = sId;
                    string name = (string)s.Attribute("name");
                    if (string.Equals(name, sheetName, StringComparison.OrdinalIgnoreCase)) {
                        sheetExists = true;
                        existingRId = (string)s.Attribute(offRelNs + "id")
                                   ?? (string)s.Attribute("r:id")
                                   ?? (string)s.Attribute("id");
                    }
                }

                if (sheetExists) {
                    foreach (var rel in relsDoc.Root.Elements(relNs + "Relationship")) {
                        if ((string)rel.Attribute("Id") == existingRId) {
                            string target = (string)rel.Attribute("Target");
                            if (!target.StartsWith("xl/")) target = "xl/" + target.TrimStart('/');
                            targetSheetPath = target;
                            break;
                        }
                    }
                    if (string.IsNullOrEmpty(targetSheetPath)) targetSheetPath = "xl/worksheets/sheet2.xml";
                    var sheetEntry = zipIn.GetEntry(targetSheetPath);
                    if (sheetEntry != null) {
                        using (var s = sheetEntry.Open()) sheetDoc = XDocument.Load(s);
                    }
                } else {
                    newSheetId = maxSheetId + 1;
                    int maxRIdNum = 0;
                    foreach (var rel in relsDoc.Root.Elements(relNs + "Relationship")) {
                        string idStr = (string)rel.Attribute("Id");
                        if (idStr != null && idStr.StartsWith("rId")) {
                            int num;
                            if (int.TryParse(idStr.Substring(3), out num) && num > maxRIdNum) maxRIdNum = num;
                        }
                    }
                    newRId = "rId" + (maxRIdNum + 1);

                    int sheetFileNum = 1;
                    while (zipIn.GetEntry("xl/worksheets/sheet" + sheetFileNum + ".xml") != null) {
                        sheetFileNum++;
                    }
                    targetRelsPath = "worksheets/sheet" + sheetFileNum + ".xml";
                    targetSheetPath = "xl/" + targetRelsPath;

                    var newSheetElem = new XElement(ns + "sheet",
                        new XAttribute("name", sheetName),
                        new XAttribute("sheetId", newSheetId),
                        new XAttribute(offRelNs + "id", newRId)
                    );
                    sheetsElem.Add(newSheetElem);

                    var newRelElem = new XElement(relNs + "Relationship",
                        new XAttribute("Id", newRId),
                        new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),
                        new XAttribute("Target", targetRelsPath)
                    );
                    relsDoc.Root.Add(newRelElem);

                    var newOverride = new XElement(ctNs + "Override",
                        new XAttribute("PartName", "/" + targetSheetPath),
                        new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
                    );
                    ctDoc.Root.Add(newOverride);

                    sheetDoc = new XDocument(
                        new XDeclaration("1.0", "UTF-8", "yes"),
                        new XElement(ns + "worksheet",
                            new XAttribute(XNamespace.Xmlns + "r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
                            new XElement(ns + "sheetData")
                        )
                    );
                }

                var sheetData = sheetDoc.Root.Element(ns + "sheetData");
                if (sheetData == null) {
                    sheetData = new XElement(ns + "sheetData");
                    sheetDoc.Root.Add(sheetData);
                }

                int currentMaxRow = 0;
                foreach (var r in sheetData.Elements(ns + "row")) {
                    int rNum = (int?)r.Attribute("r") ?? 0;
                    if (rNum > currentMaxRow) currentMaxRow = rNum;
                }

                if (currentMaxRow == 0 && headers != null && headers.Length > 0) {
                    currentMaxRow = 1;
                    var hRow = new XElement(ns + "row", new XAttribute("r", 1));
                    for (int c = 0; c < headers.Length; c++) {
                        string refStr = ColIndexToName(c) + "1";
                        var cElem = new XElement(ns + "c",
                            new XAttribute("r", refStr),
                            new XAttribute("t", "inlineStr"),
                            new XElement(ns + "is", new XElement(ns + "t", headers[c] ?? ""))
                        );
                        hRow.Add(cElem);
                    }
                    sheetData.Add(hRow);
                }

                foreach (var rowVals in rowList) {
                    currentMaxRow++;
                    var dRow = new XElement(ns + "row", new XAttribute("r", currentMaxRow));
                    for (int c = 0; c < rowVals.Count; c++) {
                        string val = rowVals[c] ?? "";
                        string refStr = ColIndexToName(c) + currentMaxRow;
                        var cElem = new XElement(ns + "c",
                            new XAttribute("r", refStr),
                            new XAttribute("t", "inlineStr"),
                            new XElement(ns + "is", new XElement(ns + "t", val))
                        );
                        if (val.Length > 0 && (char.IsWhiteSpace(val[0]) || char.IsWhiteSpace(val[val.Length - 1]))) {
                            cElem.Element(ns + "is").Element(ns + "t").SetAttributeValue(XNamespace.Xml + "space", "preserve");
                        }
                        dRow.Add(cElem);
                    }
                    sheetData.Add(dRow);
                }

                using (var outStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None))
                using (var zipOut = new ZipArchive(outStream, ZipArchiveMode.Create)) {
                    foreach (var entry in zipIn.Entries) {
                        if (string.Equals(entry.FullName, targetSheetPath, StringComparison.OrdinalIgnoreCase)) {
                            continue;
                        } else if (string.Equals(entry.FullName, "xl/workbook.xml", StringComparison.OrdinalIgnoreCase)) {
                            var e = zipOut.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                            using (var es = e.Open()) using (var xw = XmlWriter.Create(es)) wbDoc.Save(xw);
                        } else if (string.Equals(entry.FullName, "xl/_rels/workbook.xml.rels", StringComparison.OrdinalIgnoreCase)) {
                            var e = zipOut.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                            using (var es = e.Open()) using (var xw = XmlWriter.Create(es)) relsDoc.Save(xw);
                        } else if (string.Equals(entry.FullName, "[Content_Types].xml", StringComparison.OrdinalIgnoreCase)) {
                            var e = zipOut.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                            using (var es = e.Open()) using (var xw = XmlWriter.Create(es)) ctDoc.Save(xw);
                        } else {
                            var e = zipOut.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                            using (var src = entry.Open()) using (var dst = e.Open()) src.CopyTo(dst);
                        }
                    }
                    var targetEntry = zipOut.CreateEntry(targetSheetPath, CompressionLevel.Optimal);
                    using (var s = targetEntry.Open()) using (var xw = XmlWriter.Create(s)) sheetDoc.Save(xw);
                }
            }
        } catch (IOException ioEx) {
            throw new IOException("Base file is locked by another process (e.g. Microsoft Excel). Close the file and retry.\n(" + ioEx.Message + ")", ioEx);
        }

        try {
            File.Replace(tempPath, filePath, null);
        } catch {
            File.Copy(tempPath, filePath, true);
            try { File.Delete(tempPath); } catch { }
        }
    }
}
"@
    if ($PSVersionTable.PSVersion.Major -le 5) {
        Add-Type -TypeDefinition $editCSharp -ReferencedAssemblies @(
            'System.Xml',
            'System.Xml.Linq',
            'System.Core',
            'System.IO.Compression',
            'System.IO.Compression.FileSystem',
            ([PSObject].Assembly.Location)
        ) -Language CSharp
    } else {
        Add-Type -TypeDefinition $editCSharp -Language CSharp
    }
}

# ==============================================================================
# Region 2: Configuration Management (Get-AppConfig / Save-AppConfig)
# ==============================================================================
$script:DefaultAppDataDir = Join-Path $env:APPDATA 'MasterUpdater'
$script:DefaultConfigPath = Join-Path $script:DefaultAppDataDir 'config.json'

<#
.SYNOPSIS
    Returns default application configuration settings.

.DESCRIPTION
    Constructs an ordered dictionary containing default paths, write modes, backup retention counts,
    and configured metadata stamping definitions.

.OUTPUTS
    System.Collections.Specialized.OrderedDictionary.
#>
function Get-DefaultAppConfig {
    $baseDir = $script:DefaultAppDataDir
    return [ordered]@{
        ConfigVersion          = '1.0'
        BaseFilePath           = ''
        BaseSheet              = ''
        RememberBasePath       = $true
        ProfileStorePath       = (Join-Path $baseDir 'Profiles')
        LogDirectory           = (Join-Path $baseDir 'Logs')
        LogFormats             = @('jsonl', 'txt')
        BackupDirectory        = (Join-Path $baseDir 'Backups')
        BackupRetentionCount   = 20
        MaxBackupMb            = 200    # P6: Auto-prune backup folder when total size exceeds this threshold (MB). 0 = disabled.
        WriteMode              = 'InPlace' # 'InPlace' or 'SafeRewrite'
        RedactNamesInLog       = $false
        DetectRemovedRows      = $false
        AutoSkipUnchanged      = $true
        ShowUnchangedRows      = $false
        LogChangesToBaseSheet  = $false
        BaseSheetLogName       = 'ImportLog'
        MarkDeletedMode        = 'ColumnStatus' # 'ColumnStatus' or 'PhysicallyDelete'
        MarkDeletedColumn      = 'StatusOpieki'
        MarkDeletedValue       = 'Usunięty'
        CompareIgnoreCase      = $true
        CompareTrimWhitespace  = $true
        CompareIgnoreSpecialChars = $true
        CompareIgnoreAllSpaces = $true
        UiLanguage             = 'PL' # 'PL', 'EN', 'DE'
        Theme                  = 'Dark'
        MetadataColumns        = @(
            [ordered]@{ BaseColumn = 'OstZmiana';      Token = 'ChangeDate';         Format = 'yyyy-MM-dd HH:mm' },
            [ordered]@{ BaseColumn = 'ZrodloSciezka';  Token = 'SourceFileFullPath'; Format = '' },
            [ordered]@{ BaseColumn = 'ZrodloPlik';     Token = 'SourceFileName';     Format = '' },
            [ordered]@{ BaseColumn = 'Zmienil';        Token = 'CurrentUser';        Format = '' },
            [ordered]@{ BaseColumn = 'ZrodloWiersz';   Token = 'SourceRowNumber';    Format = '' },
            [ordered]@{ BaseColumn = 'ImportId';       Token = 'ImportBatchId';      Format = '' }
        )
    }
}

<#
.SYNOPSIS
    Loads application configuration from disk.

.DESCRIPTION
    Reads configuration JSON from the specified path. Merges saved properties with default
    settings, creating default configuration if the file does not exist.

.PARAMETER Path
    Path to configuration JSON file. Defaults to %APPDATA%\MasterUpdater\config.json.

.OUTPUTS
    System.Collections.Specialized.OrderedDictionary.
#>
function Get-AppConfig {
    [CmdletBinding()]
    param(
        [Alias('ConfigPath')]
        [string]$Path = $script:DefaultConfigPath
    )
    if ([string]::IsNullOrWhiteSpace($Path)) {
        $Path = $script:DefaultConfigPath
    }
    $default = Get-DefaultAppConfig
    if (-not (Test-Path $Path)) {
        Save-AppConfig -Config $default -Path $Path
        return $default
    }
    try {
        $raw = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
        $loaded = $raw | ConvertFrom-Json
        $cfg = Get-DefaultAppConfig
        foreach ($prop in $loaded.PSObject.Properties) {
            $cfg[$prop.Name] = $prop.Value
        }
        if ($loaded.PSObject.Properties['ShowUnchangedRows']) {
            $cfg['AutoSkipUnchanged'] = ($cfg['ShowUnchangedRows'] -eq $false)
        } elseif ($loaded.PSObject.Properties['AutoSkipUnchanged']) {
            $cfg['ShowUnchangedRows'] = ($cfg['AutoSkipUnchanged'] -eq $false)
        }
        if ($loaded.PSObject.Properties['LogChangesToBaseSheet']) {
            $cfg['LogChangesToBaseSheet'] = [bool]$loaded.LogChangesToBaseSheet
        }
        if ($loaded.PSObject.Properties['BaseSheetLogName'] -and -not [string]::IsNullOrWhiteSpace($loaded.BaseSheetLogName)) {
            $cfg['BaseSheetLogName'] = [string]$loaded.BaseSheetLogName
        }
        return $cfg
    } catch {
        return $default
    }
}

<#
.SYNOPSIS
    Saves application configuration to disk.

.DESCRIPTION
    Serializes the configuration hashtable to formatted JSON and writes it using UTF-8 with BOM.

.PARAMETER Config
    Hashtable of application configuration options.

.PARAMETER Path
    Destination file path.

.OUTPUTS
    System.Void.
#>
function Save-AppConfig {
    param([hashtable]$Config, [string]$Path = $script:DefaultConfigPath)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        $Path = $script:DefaultConfigPath
    }
    $dir = Split-Path $Path -Parent
    if (-not [string]::IsNullOrEmpty($dir) -and -not (Test-Path $dir)) {
        [void][System.IO.Directory]::CreateDirectory($dir)
    }
    $json = $Config | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($Path, $json, [System.Text.UTF8Encoding]::new($true))
}

# ==============================================================================
# Region 3: Profile Store & Header Fingerprint Auto-Resolution
# ==============================================================================
<#
.SYNOPSIS
    Computes deterministic SHA-256 fingerprint for worksheet headers.

.DESCRIPTION
    Normalizes header names (trimmed, lowercased) and worksheet name, concatenating them
    into a pipe-separated string and hashing with SHA-256 to produce a unique template signature.

.PARAMETER Headers
    Array of column header names.

.PARAMETER SheetName
    Optional worksheet name.

.OUTPUTS
    System.String. 64-character uppercase hexadecimal hash.
#>
Set-Alias -Name 'Get-HeaderFingerprint' -Value 'Compute-HeaderFingerprint' -ErrorAction SilentlyContinue
function Compute-HeaderFingerprint {
    param([string[]]$Headers, [string]$SheetName = '')
    $prefix = if ([string]::IsNullOrWhiteSpace($SheetName)) { '' } else { $SheetName.Trim().ToLowerInvariant() + '#' }
    $normList = @($Headers | ForEach-Object { $_.Trim().ToLowerInvariant() })
    $raw = $prefix + ($normList -join '|')
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($raw)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = $sha.ComputeHash($bytes)
        return -join ($hash | ForEach-Object { $_.ToString('X2') })
    } finally {
        $sha.Dispose()
    }
}

<#
.SYNOPSIS
    Finds a saved mapping profile matching a header fingerprint.

.DESCRIPTION
    Iterates over JSON profiles in the profile store directory and returns the first profile
    matching the specified SHA-256 header fingerprint.

.PARAMETER Fingerprint
    Target 64-character SHA-256 fingerprint.

.PARAMETER ProfileDir
    Directory path containing saved profile JSON files.

.OUTPUTS
    System.Collections.Hashtable or $null if no match is found.
#>
function Get-MappingProfileByFingerprint {
    param([string]$Fingerprint, [string]$StorePath)
    if (-not (Test-Path $StorePath)) { return $null }
    $profileFiles = Get-ChildItem -Path $StorePath -Filter '*.json' -File
    foreach ($file in $profileFiles) {
        try {
            $raw = [System.IO.File]::ReadAllText($file.FullName, [System.Text.Encoding]::UTF8)
            $p = $raw | ConvertFrom-Json
            if ($p.HeaderFingerprint -eq $Fingerprint) {
                return $p
            }
        } catch { }
    }
    return $null
}

<#
.SYNOPSIS
    Saves a column mapping profile to disk.

.DESCRIPTION
    Saves mapping rules, join keys, and header fingerprint to an individual JSON file in the profile store.

.PARAMETER Profile
    Hashtable containing profile schema, name, rules, and fingerprint.

.PARAMETER ProfileDir
    Target directory for mapping profile storage.

.OUTPUTS
    System.String. Full path of saved profile file.
#>
function Save-MappingProfile {
    param(
        [object]$Profile,
        [string]$StorePath
    )
    if (-not (Test-Path $StorePath)) {
        [void][System.IO.Directory]::CreateDirectory($StorePath)
    }
    $safeName = ($Profile.Name -replace '[^\w\.\-]+', '_').Trim('_')
    if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = 'Profile_' + (Get-Date -Format 'yyyyMMdd_HHmmss') }
    $filePath = Join-Path $StorePath "$safeName.json"
    $json = $Profile | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($filePath, $json, [System.Text.UTF8Encoding]::new($true))
    return $filePath
}

<#
.SYNOPSIS
    Returns localized display name for a mapping rule DiffPolicy.
#>
function Get-DiffPolicyDisplay {
    param([string]$Policy)
    switch ($Policy) {
        'IgnoreChanges'       { return (Get-UiString 'DiffPolicyIgnoreChanges' 'Ignoruj zmiany (tylko nowe wpisy)') }
        'NormalizePostalCode' { return (Get-UiString 'DiffPolicyNormalizePostal' 'Ignoruj obecność kodu pocztowego') }
        'FuzzyContainment'    { return (Get-UiString 'DiffPolicyFuzzyContainment' 'Ignoruj dopiski w nazwach (zawieranie tekstu)') }
        default               { return (Get-UiString 'DiffPolicyTrackChanges' 'Śledź zmiany (standard)') }
    }
}

<#
.SYNOPSIS
    Smart fuzzy matching between incoming column headers and base file headers.
#>
function Find-SmartHeaderMatch {
    param(
        [string]$IncomingHeader,
        [string[]]$BaseHeaders
    )
    if ([string]::IsNullOrWhiteSpace($IncomingHeader) -or -not $BaseHeaders) { return $null }
    $cleanIh = ($IncomingHeader -replace '[^\p{L}\p{Nd}\s]', ' ').Trim().ToLowerInvariant()
    $normIh = $cleanIh -replace '\s+', ''
    $ihWords = @($cleanIh -split '\s+' | Where-Object { $_.Length -gt 2 })

    # 1. Exact or substring match (ignoring special chars & whitespace)
    foreach ($bh in $BaseHeaders) {
        if ([string]::IsNullOrWhiteSpace($bh)) { continue }
        $cleanBh = ($bh -replace '[^\p{L}\p{Nd}\s]', ' ').Trim().ToLowerInvariant()
        $normBh = $cleanBh -replace '\s+', ''
        if ($normBh -eq $normIh -or $normBh.Contains($normIh) -or $normIh.Contains($normBh)) {
            return $bh
        }
    }

    # 2. Word overlap / semantic stem match (>= 50% words shared)
    if ($ihWords.Count -ge 2) {
        $bestMatch = $null
        $bestScore = 0.0
        foreach ($bh in $BaseHeaders) {
            if ([string]::IsNullOrWhiteSpace($bh)) { continue }
            $cleanBh = ($bh -replace '[^\p{L}\p{Nd}\s]', ' ').Trim().ToLowerInvariant()
            $bhWords = @($cleanBh -split '\s+' | Where-Object { $_.Length -gt 2 })
            if ($bhWords.Count -eq 0) { continue }

            $shared = 0
            foreach ($w in $ihWords) {
                $stem = if ($w.Length -ge 4) { $w.Substring(0, 4) } else { $w }
                $found = $false
                foreach ($bw in $bhWords) {
                    $bStem = if ($bw.Length -ge 4) { $bw.Substring(0, 4) } else { $bw }
                    if ($stem -eq $bStem) { $found = $true; break }
                }
                if ($found) { $shared++ }
            }
            $score = [double]$shared / [Math]::Max($ihWords.Count, $bhWords.Count)
            if ($score -gt $bestScore -and $score -ge 0.5) {
                $bestScore = $score
                $bestMatch = $bh
            }
        }
        if ($bestMatch) { return $bestMatch }
    }

    return $null
}

<#
.SYNOPSIS
    Generates smart mapping rules between base and incoming column headers.
#>
function Invoke-AutoMapRules {
    param(
        [string[]]$BaseHeaders,
        [string[]]$IncomingHeaders
    )
    $rules = [System.Collections.Generic.List[object]]::new()
    if (-not $BaseHeaders -or -not $IncomingHeaders) { return $rules }

    $incHasCity = ($IncomingHeaders | Where-Object { $_ -match '(?i)(miasto|miejscowo|city|ort|stadt)' })
    $baseCityCol = if (-not $incHasCity) {
        $BaseHeaders | Where-Object { $_ -match '(?i)(miasto|miejscowo|city|ort|stadt)' } | Select-Object -First 1
    } else { $null }

    foreach ($ih in $IncomingHeaders) {
        if ($ih -match '(?i)^l\.?\s*p\.?$') { continue }
        $matched = Find-SmartHeaderMatch -IncomingHeader $ih -BaseHeaders $BaseHeaders
        if ($matched) {
            if ($baseCityCol -and $ih -match '(?i)(adres.*zamieszk|adres.*domu|adres.*klient|adres.*dzieck|street|address)' -and $matched -notmatch '(?i)(pracy|firm|szko|plac|zak|work|school)') {
                $rules.Add([PSCustomObject]@{
                    BaseColumns       = @($matched, $baseCityCol)
                    UpdateColumns     = @($ih)
                    MergeMode         = 'Concatenate'
                    Separator         = ', '
                    BaseColsStr       = "$matched, $baseCityCol"
                    UpdColsStr        = $ih
                    DiffPolicy        = 'TrackChanges'
                    DiffPolicyDisplay = (Get-DiffPolicyDisplay 'TrackChanges')
                })
            } else {
                $rules.Add([PSCustomObject]@{
                    BaseColumns       = @($matched)
                    UpdateColumns     = @($ih)
                    MergeMode         = 'Exact'
                    Separator         = ''
                    BaseColsStr       = $matched
                    UpdColsStr        = $ih
                    DiffPolicy        = 'TrackChanges'
                    DiffPolicyDisplay = (Get-DiffPolicyDisplay 'TrackChanges')
                })
            }
        }
    }
    return $rules
}

# ==============================================================================
# Region 4: Import Pipeline & Mapping Projection (Get-ProjectedRow)
# ==============================================================================
<#
.SYNOPSIS
    Projects incoming row values into master base-column space.

.DESCRIPTION
    Applies column mapping rules to translate incoming source fields to base fields,
    handling N:1 merges (Exact, FirstNonEmpty, Concatenate) and 1:N splits (delimiter or regex).

.PARAMETER IncomingValues
    Hashtable of source row column values.

.PARAMETER MappingRules
    List of mapping rule definitions.

.OUTPUTS
    System.Collections.Hashtable of projected base column values.
#>
function Get-ProjectedRow {
    param(
        [Parameter(Mandatory = $false)]
        [Alias('IncomingRow')]
        $IncomingValues,
        [Parameter(Mandatory = $false)]
        [object[]]$MappingRules,
        [bool]$Trim = $true
    )
    $valDict = @{}
    if ($IncomingValues -is [System.Collections.IDictionary]) {
        foreach ($k in $IncomingValues.Keys) {
            $valDict[$k] = $IncomingValues[$k]
        }
    } elseif ($IncomingValues -is [System.Data.DataRow]) {
        foreach ($col in $IncomingValues.Table.Columns) {
            $valDict[$col.ColumnName] = if ($IncomingValues.IsNull($col)) { '' } else { $IncomingValues[$col].ToString() }
        }
    } elseif ($null -ne $IncomingValues) {
        foreach ($p in $IncomingValues.PSObject.Properties) {
            $valDict[$p.Name] = if ($null -ne $p.Value) { $p.Value.ToString() } else { '' }
        }
    }

    $projected = @{}
    foreach ($rule in $MappingRules) {
        $bCols = if ($rule -is [System.Collections.IDictionary]) { if ($rule.Contains('BaseColumns')) { $rule['BaseColumns'] } else { $null } } elseif ($rule.PSObject.Properties['BaseColumns']) { $rule.BaseColumns } else { $null }
        $uCols = if ($rule -is [System.Collections.IDictionary]) { if ($rule.Contains('UpdateColumns')) { $rule['UpdateColumns'] } else { $null } } elseif ($rule.PSObject.Properties['UpdateColumns']) { $rule.UpdateColumns } else { $null }
        $mMode = if ($rule -is [System.Collections.IDictionary]) { if ($rule.Contains('MergeMode')) { $rule['MergeMode'] } else { $null } } elseif ($rule.PSObject.Properties['MergeMode']) { $rule.MergeMode } else { $null }
        $rSep  = if ($rule -is [System.Collections.IDictionary]) { if ($rule.Contains('Separator')) { $rule['Separator'] } else { $null } } elseif ($rule.PSObject.Properties['Separator']) { $rule.Separator } else { $null }
        $sRegex = if ($rule -is [System.Collections.IDictionary]) { if ($rule.Contains('SplitRegex')) { $rule['SplitRegex'] } else { $null } } elseif ($rule.PSObject.Properties['SplitRegex']) { $rule.SplitRegex } else { $null }
        $sSep   = if ($rule -is [System.Collections.IDictionary]) { if ($rule.Contains('SplitSeparator')) { $rule['SplitSeparator'] } else { $null } } elseif ($rule.PSObject.Properties['SplitSeparator']) { $rule.SplitSeparator } else { $null }

        $baseCols = @($bCols | Where-Object { $null -ne $_ -and [string]::IsNullOrWhiteSpace($_) -eq $false })
        $updCols  = @($uCols | Where-Object { $null -ne $_ -and [string]::IsNullOrWhiteSpace($_) -eq $false })
        $mode     = if (-not [string]::IsNullOrEmpty($mMode)) { $mMode } else { 'Exact' }
        $sep      = if ($null -ne $rSep) { $rSep } else { '' }

        if ($baseCols.Length -eq 0 -or $updCols.Length -eq 0) {
            continue
        }

        # Case 1: Regex split (1 update col -> multiple base cols)
        if (-not [string]::IsNullOrEmpty($sRegex) -and $updCols.Length -ge 1) {
            $srcVal = if (-not [string]::IsNullOrEmpty($updCols[0]) -and $valDict.ContainsKey($updCols[0])) { $valDict[$updCols[0]] } else { '' }
            if ($srcVal -match $sRegex) {
                for ($b = 0; $b -lt $baseCols.Length; $b++) {
                    $bName = $baseCols[$b]
                    $val = if ($Matches.ContainsKey($bName)) {
                        $Matches[$bName]
                    } elseif ($Matches.ContainsKey(($b + 1).ToString())) {
                        $Matches[($b + 1).ToString()]
                    } else { '' }
                    $projected[$bName] = if ($null -ne $val) { if ($Trim) { $val.ToString().Trim() } else { $val.ToString() } } else { '' }
                }
            }
            continue
        }

        # Case 2: Delimiter split (1 update col -> multiple base cols)
        $sepToSplit = if (-not [string]::IsNullOrEmpty($sSep)) { $sSep } elseif ($mode -eq 'Concatenate' -and -not [string]::IsNullOrEmpty($sep)) { $sep } else { '' }
        if (-not [string]::IsNullOrEmpty($sepToSplit) -and $updCols.Length -ge 1 -and $baseCols.Length -gt 1) {
            $srcVal = if (-not [string]::IsNullOrEmpty($updCols[0]) -and $valDict.ContainsKey($updCols[0])) { $valDict[$updCols[0]] } else { '' }
            $parts = $srcVal -split [regex]::Escape($sepToSplit)
            for ($b = 0; $b -lt $baseCols.Length; $b++) {
                $bName = $baseCols[$b]
                $val = if ($b -lt $parts.Length) { if ($Trim) { $parts[$b].Trim() } else { $parts[$b] } } else { '' }
                $projected[$bName] = $val
            }
            continue
        }

        # Case 3: N-to-1 merge (N update cols -> 1 base col) or 1-to-1 exact
        if ($baseCols.Length -ge 1) {
            $vals = [System.Collections.Generic.List[string]]::new()
            foreach ($u in $updCols) {
                if (-not [string]::IsNullOrEmpty($u) -and $valDict.ContainsKey($u) -and $null -ne $valDict[$u]) {
                    $vals.Add($valDict[$u].ToString())
                } else {
                    $vals.Add('')
                }
            }
            $merged = [FastDiffHelper]::MergeValues($vals, $mode, $sep, $Trim)
            $projected[$baseCols[0]] = $merged
        }
    }
    return $projected
}

<#
.SYNOPSIS
    Generates a normalized composite join key from specified row columns.

.DESCRIPTION
    Extracts key values from row hashtable, normalizes strings according to comparison options
    (case, whitespace, special characters), and joins them with pipe separator.

.PARAMETER RowValues
    Hashtable of row values.

.PARAMETER KeyColumns
    Array of column names forming the composite key.

.PARAMETER Trim
    Whether to trim leading/trailing whitespace.

.PARAMETER IgnoreCase
    Whether to lowercase key values.

.PARAMETER IgnoreSpecialChars
    Whether to strip non-alphanumeric characters.

.PARAMETER IgnoreAllSpaces
    Whether to collapse or remove whitespace.

.OUTPUTS
    System.String. Normalized composite join key.
#>
function Build-JoinKey {
    param(
        [hashtable]$RowValues,
        [string[]]$KeyColumns,
        [bool]$Trim = $true,
        [bool]$IgnoreCase = $true,
        [bool]$IgnoreSpecialChars = $true,
        [bool]$IgnoreAllSpaces = $true
    )
    $sb = [System.Text.StringBuilder]::new()
    for ($i = 0; $i -lt $KeyColumns.Length; $i++) {
        $col = $KeyColumns[$i]
        $v = if ($RowValues -is [System.Collections.IDictionary]) {
            if ($RowValues.Contains($col) -and $null -ne $RowValues[$col]) { $RowValues[$col].ToString() } else { '' }
        } elseif ($RowValues -and $RowValues.PSObject.Properties[$col]) {
            if ($null -ne $RowValues.$col) { $RowValues.$col.ToString() } else { '' }
        } else { '' }
        $norm = [FastDiffHelper]::Normalize($v, $IgnoreCase, $Trim, $IgnoreSpecialChars, $IgnoreAllSpaces)
        if ($i -gt 0) { [void]$sb.Append('|||') }
        [void]$sb.Append($norm)
    }
    return $sb.ToString()
}

# ==============================================================================
# Region 5: Compare Engine (Invoke-MasterCompare)
# ==============================================================================
<#
.SYNOPSIS
    Compares incoming rows against master base rows and produces review items.

.DESCRIPTION
    Performs full join-key matching, detects New rows, Changed cells with exact coordinates (CellRef),
    Ambiguous duplicate or missing keys, and Removed rows when enabled.

.PARAMETER BaseRows
    List of base rows with RowNumber and Values.

.PARAMETER IncomingRows
    List of incoming rows with SourceRowNumber and Values.

.PARAMETER MappingProfile
    Mapping profile object.

.PARAMETER MappingRules
    Optional list of mapping rules.

.PARAMETER BaseJoinKey
    Base join key column array.

.PARAMETER IncomingJoinKey
    Incoming join key column array.

.PARAMETER BaseHeaders
    Array of base file headers.

.PARAMETER MetadataColumns
    Array of metadata columns to exclude from comparison.

.PARAMETER DetectRemoved
    Whether to detect missing base rows as Removed.

.PARAMETER MarkDeletedColumn
    Base column to update for removed rows (e.g. StatusOpieki).

.PARAMETER MarkDeletedValue
    Value to set for removed rows (e.g. Usunięty).

.OUTPUTS
    System.Collections.Generic.List[PSCustomObject] containing review items.
#>
function Invoke-MasterCompare {
    param(
        [System.Collections.IList]$BaseRows,         # List of PSCustomObject @{ RowNumber; Values }
        [System.Collections.IList]$IncomingRows,     # List of PSCustomObject @{ SourceRowNumber; Values; SourceFilePath }
        [object]$MappingProfile,
        [object[]]$MappingRules,
        [string[]]$BaseJoinKey,
        [string[]]$IncomingJoinKey,
        [string[]]$BaseHeaders,
        [hashtable]$CompareOptions = $null,
        [string[]]$MetadataColumns = @('OstZmiana', 'ZrodloSciezka', 'ZrodloPlik', 'Zmienil', 'ZrodloWiersz', 'ImportId'),
        [bool]$DetectRemoved = $false,
        [string]$MarkDeletedColumn = 'StatusOpieki',
        [string]$MarkDeletedValue = 'Usunięty'
    )

    $rules = if ($MappingRules) { @($MappingRules) } elseif ($MappingProfile -and $MappingProfile.MappingRules) { @($MappingProfile.MappingRules) } else { @() }
    $joinBase = if ($BaseJoinKey) { @($BaseJoinKey) } elseif ($MappingProfile -and $MappingProfile.JoinKeyBase) { @($MappingProfile.JoinKeyBase) } else { @() }
    $joinUpd = if ($IncomingJoinKey) { @($IncomingJoinKey) } elseif ($MappingProfile -and $MappingProfile.JoinKeyUpdate) { @($MappingProfile.JoinKeyUpdate) } else { @() }
    $rawOpts = if ($CompareOptions) {
        $CompareOptions
    } elseif ($MappingProfile -and $MappingProfile.CompareOptions) {
        $MappingProfile.CompareOptions
    } else {
        @{ IgnoreCase = $true; Trim = $true; IgnoreSpecialChars = $true; IgnoreAllSpaces = $true }
    }
    $optTrim = if ($rawOpts -is [System.Collections.IDictionary]) {
        if ($rawOpts.Contains('Trim')) { [bool]$rawOpts['Trim'] } elseif ($rawOpts.Contains('TrimWhitespace')) { [bool]$rawOpts['TrimWhitespace'] } else { $true }
    } elseif ($rawOpts.PSObject.Properties['Trim']) {
        [bool]$rawOpts.Trim
    } elseif ($rawOpts.PSObject.Properties['TrimWhitespace']) {
        [bool]$rawOpts.TrimWhitespace
    } else { $true }

    $optCase = if ($rawOpts -is [System.Collections.IDictionary]) {
        if ($rawOpts.Contains('IgnoreCase')) { [bool]$rawOpts['IgnoreCase'] } else { $true }
    } elseif ($rawOpts.PSObject.Properties['IgnoreCase']) {
        [bool]$rawOpts.IgnoreCase
    } else { $true }

    $optSpec = if ($rawOpts -is [System.Collections.IDictionary]) {
        if ($rawOpts.Contains('IgnoreSpecialChars')) { [bool]$rawOpts['IgnoreSpecialChars'] } else { $true }
    } elseif ($rawOpts.PSObject.Properties['IgnoreSpecialChars']) {
        [bool]$rawOpts.IgnoreSpecialChars
    } else { $true }

    $optSpc = if ($rawOpts -is [System.Collections.IDictionary]) {
        if ($rawOpts.Contains('IgnoreAllSpaces')) { [bool]$rawOpts['IgnoreAllSpaces'] } else { $true }
    } elseif ($rawOpts.PSObject.Properties['IgnoreAllSpaces']) {
        [bool]$rawOpts.IgnoreAllSpaces
    } else { $true }

    $opts = @{
        IgnoreCase         = $optCase
        Trim               = $optTrim
        TrimWhitespace     = $optTrim
        IgnoreSpecialChars = $optSpec
        IgnoreAllSpaces    = $optSpc
    }
    $emptyMeansClear = if ($MappingProfile) { [bool]$MappingProfile.EmptyIncomingMeansClear } else { $false }

    # Normalize BaseRows: ensure each item has .RowNumber and .Values
    $normalizedBaseRows = [System.Collections.Generic.List[object]]::new()
    foreach ($b in $BaseRows) {
        if ($b.PSObject.Properties['Values']) {
            if ($b.Values -is [System.Collections.IDictionary] -and $b.Values -isnot [System.Collections.Hashtable]) {
                $h = @{}
                foreach ($k in $b.Values.Keys) { $h[$k] = $b.Values[$k] }
                $b.Values = $h
            }
            $normalizedBaseRows.Add($b)
        } else {
            $vals = @{}
            foreach ($p in $b.PSObject.Properties) {
                if ($p.Name -ne '_RowNumber') { $vals[$p.Name] = $p.Value }
            }
            $rNum = if ($b.PSObject.Properties['_RowNumber'] -and $b._RowNumber) { [int]$b._RowNumber } else { ($normalizedBaseRows.Count + 2) }
            $normalizedBaseRows.Add([PSCustomObject]@{ RowNumber = $rNum; Values = $vals })
        }
    }

    # Normalize IncomingRows: ensure each item has .SourceRowNumber and .Values
    $normalizedIncRows = [System.Collections.Generic.List[object]]::new()
    foreach ($inc in $IncomingRows) {
        if ($inc.PSObject.Properties['Values']) {
            if ($inc.Values -is [System.Collections.IDictionary] -and $inc.Values -isnot [System.Collections.Hashtable]) {
                $h = @{}
                foreach ($k in $inc.Values.Keys) { $h[$k] = $inc.Values[$k] }
                $inc.Values = $h
            }
            # Ensure SourceFilePath is present even for pre-normalized rows
            if (-not $inc.PSObject.Properties['SourceFilePath']) {
                $inc.PSObject.Properties.Add([System.Management.Automation.PSNoteProperty]::new('SourceFilePath', ''))
            }
            $normalizedIncRows.Add($inc)
        } else {
            $vals = @{}
            foreach ($p in $inc.PSObject.Properties) {
                if ($p.Name -ne '_RowNumber') { $vals[$p.Name] = $p.Value }
            }
            $rNum = if ($inc.PSObject.Properties['_RowNumber'] -and $inc._RowNumber) { [int]$inc._RowNumber } else { ($normalizedIncRows.Count + 2) }
            # Preserve SourceFilePath from source object if it was stamped before normalization
            $srcFilePath = if ($inc.PSObject.Properties['SourceFilePath'] -and -not [string]::IsNullOrEmpty($inc.SourceFilePath)) { $inc.SourceFilePath } else { '' }
            $normalizedIncRows.Add([PSCustomObject]@{ SourceRowNumber = $rNum; Values = $vals; SourceFilePath = $srcFilePath })
        }
    }

    # 1. Index base rows by join key
    $baseIndex = @{}
    foreach ($b in $normalizedBaseRows) {
        $k = Build-JoinKey -RowValues $b.Values -KeyColumns $joinBase -Trim $opts.Trim -IgnoreCase $opts.IgnoreCase -IgnoreSpecialChars $opts.IgnoreSpecialChars -IgnoreAllSpaces $opts.IgnoreAllSpaces
        if (-not $baseIndex.ContainsKey($k)) {
            $baseIndex[$k] = [System.Collections.Generic.List[object]]::new()
        }
        $baseIndex[$k].Add($b)
    }

    # 2. Iterate incoming rows
    $reviewItems = [System.Collections.Generic.List[object]]::new()
    $incomingKeysSeen = @{}
    $claimedBaseRows = [System.Collections.Generic.HashSet[int]]::new()
    $idx = 0

    foreach ($inc in $normalizedIncRows) {
        $idx++
        $srcRowNum = $inc.SourceRowNumber
        $projected = Get-ProjectedRow -IncomingValues $inc.Values -MappingRules $rules -Trim $opts.Trim

        # Check for empty incoming join key
        $hasEmptyKey = $false
        foreach ($kCol in $joinUpd) {
            $rawK = if ($inc.Values.ContainsKey($kCol)) { $inc.Values[$kCol] } else { '' }
            if ([string]::IsNullOrWhiteSpace($rawK)) { $hasEmptyKey = $true; break }
        }

        if ($hasEmptyKey) {
            $reviewItems.Add([PSCustomObject]@{
                Index             = $idx
                Status            = 'Ambiguous'
                Decision          = 'Pending'
                AmbiguousReason   = (Get-UiString 'AmbiguousReasonEmptyKey' 'Pusty klucz złączenia w pliku źródłowym')
                IncomingRow       = $inc
                ProjectedRow      = $projected
                MatchedBaseRow    = $null
                CandidateBaseRows = @()
                Changes           = @()
                DecisionBy        = $env:USERNAME
                DecisionUtc       = (Get-Date).ToUniversalTime().ToString('o')
            })
            continue
        }

        $uKey = Build-JoinKey -RowValues $inc.Values -KeyColumns $joinUpd -Trim $opts.Trim -IgnoreCase $opts.IgnoreCase -IgnoreSpecialChars $opts.IgnoreSpecialChars -IgnoreAllSpaces $opts.IgnoreAllSpaces

        # Check duplicate incoming key
        if ($incomingKeysSeen.ContainsKey($uKey)) {
            $reviewItems.Add([PSCustomObject]@{
                Index             = $idx
                Status            = 'Ambiguous'
                Decision          = 'Pending'
                AmbiguousReason   = ((Get-UiString 'AmbiguousReasonDupKey' "Zduplikowany klucz w pliku źródłowym ('{0}')") -f $uKey)
                IncomingRow       = $inc
                ProjectedRow      = $projected
                MatchedBaseRow    = $null
                CandidateBaseRows = @()
                Changes           = @()
                DecisionBy        = $env:USERNAME
                DecisionUtc       = (Get-Date).ToUniversalTime().ToString('o')
            })
            continue
        }
        $incomingKeysSeen[$uKey] = $srcRowNum

        # Match in base
        $matchList = @(if ($baseIndex.ContainsKey($uKey)) { $baseIndex[$uKey] } else { })

        if ($matchList.Count -eq 0) {
            # New row
            $reviewItems.Add([PSCustomObject]@{
                Index             = $idx
                Status            = 'New'
                Decision          = 'Pending'
                AmbiguousReason   = $null
                IncomingRow       = $inc
                ProjectedRow      = $projected
                MatchedBaseRow    = $null
                CandidateBaseRows = @()
                Changes           = @()
                DecisionBy        = $env:USERNAME
                DecisionUtc       = (Get-Date).ToUniversalTime().ToString('o')
            })
        } elseif ($matchList.Count -gt 1) {
            # Ambiguous - multiple matches
            $reviewItems.Add([PSCustomObject]@{
                Index             = $idx
                Status            = 'Ambiguous'
                Decision          = 'Pending'
                AmbiguousReason   = ((Get-UiString 'AmbiguousReasonMultiMatch' "Znaleziono {0} pasujących wierszy w bazie") -f $matchList.Count)
                IncomingRow       = $inc
                ProjectedRow      = $projected
                MatchedBaseRow    = $null
                CandidateBaseRows = @($matchList)
                Changes           = @()
                DecisionBy        = $env:USERNAME
                DecisionUtc       = (Get-Date).ToUniversalTime().ToString('o')
            })
        } else {
            # Exactly 1 match
            $matched = $matchList[0]
            if ($claimedBaseRows.Contains($matched.RowNumber)) {
                # Base row already claimed by earlier incoming row
                $reviewItems.Add([PSCustomObject]@{
                    Index             = $idx
                    Status            = 'Ambiguous'
                    Decision          = 'Pending'
                    AmbiguousReason   = ((Get-UiString 'AmbiguousReasonAlreadyClaimed' "Wiersz bazy #{0} został już wcześniej powiązany w tej sesji") -f $matched.RowNumber)
                    IncomingRow       = $inc
                    ProjectedRow      = $projected
                    MatchedBaseRow    = $matched
                    CandidateBaseRows = @($matched)
                    Changes           = @()
                    DecisionBy        = $env:USERNAME
                    DecisionUtc       = (Get-Date).ToUniversalTime().ToString('o')
                })
            } else {
                [void]$claimedBaseRows.Add($matched.RowNumber)
                $changes = [System.Collections.Generic.List[object]]::new()

                # Compare all mapped columns
                foreach ($rule in $rules) {
                    $bCols = if ($rule -is [System.Collections.IDictionary]) { if ($rule.Contains('BaseColumns')) { @($rule['BaseColumns']) } else { @() } } elseif ($rule.PSObject.Properties['BaseColumns']) { @($rule.BaseColumns) } else { @() }
                    $uCols = if ($rule -is [System.Collections.IDictionary]) { if ($rule.Contains('UpdateColumns')) { @($rule['UpdateColumns']) } else { @() } } elseif ($rule.PSObject.Properties['UpdateColumns']) { @($rule.UpdateColumns) } else { @() }
                    $mMode = if ($rule -is [System.Collections.IDictionary]) { if ($rule.Contains('MergeMode')) { $rule['MergeMode'] } else { 'Exact' } } elseif ($rule.PSObject.Properties['MergeMode']) { $rule.MergeMode } else { 'Exact' }
                    $sep   = if ($rule -is [System.Collections.IDictionary]) { if ($rule.Contains('Separator')) { $rule['Separator'] } else { '' } } elseif ($rule.PSObject.Properties['Separator']) { $rule.Separator } else { '' }
                    $diffPolicy = if ($rule -is [System.Collections.IDictionary]) { if ($rule.Contains('DiffPolicy')) { $rule['DiffPolicy'] } else { 'TrackChanges' } } elseif ($rule.PSObject.Properties['DiffPolicy']) { $rule.DiffPolicy } else { 'TrackChanges' }
                    if ([string]::IsNullOrWhiteSpace($diffPolicy)) { $diffPolicy = 'TrackChanges' }

                    # Skip change generation for existing rows if policy is IgnoreChanges (values stay projected for new rows)
                    if ($diffPolicy -eq 'IgnoreChanges') {
                        continue
                    }

                    $bCols = @($bCols | Where-Object { $null -ne $_ -and [string]::IsNullOrWhiteSpace($_) -eq $false })
                    $uCols = @($uCols | Where-Object { $null -ne $_ -and [string]::IsNullOrWhiteSpace($_) -eq $false })

                    # Multi-column Concatenate equality check (e.g. 2 Base columns <-> 1 Incoming column, or vice versa)
                    if ($mMode -eq 'Concatenate' -and ($bCols.Length -gt 1 -or $uCols.Length -gt 1)) {
                        $bVals = [System.Collections.Generic.List[string]]::new()
                        foreach ($bc in $bCols) {
                            $v = if ($matched.Values.ContainsKey($bc) -and $null -ne $matched.Values[$bc]) { $matched.Values[$bc].ToString() } else { '' }
                            $bVals.Add($v)
                        }
                        $bMerged = [FastDiffHelper]::MergeValues($bVals, $mMode, $sep, $opts.Trim)

                        $uVals = [System.Collections.Generic.List[string]]::new()
                        foreach ($uc in $uCols) {
                            $v = if ($inc.Values.ContainsKey($uc) -and $null -ne $inc.Values[$uc]) { $inc.Values[$uc].ToString() } else { '' }
                            $uVals.Add($v)
                        }
                        $uMerged = [FastDiffHelper]::MergeValues($uVals, $mMode, $sep, $opts.Trim)

                        $isEqualOverall = [FastDiffHelper]::AreEqual($bMerged, $uMerged, $opts.IgnoreCase, $opts.Trim, $opts.IgnoreSpecialChars, $opts.IgnoreAllSpaces)
                        if (-not $isEqualOverall -and $diffPolicy -eq 'NormalizePostalCode') {
                            $bmNorm = ($bMerged -replace '\b\d{2}-\d{3}\b|\b\d{5}\b', ' ') -replace '\s+', ' '
                            $umNorm = ($uMerged -replace '\b\d{2}-\d{3}\b|\b\d{5}\b', ' ') -replace '\s+', ' '
                            $isEqualOverall = [FastDiffHelper]::AreEqual($bmNorm, $umNorm, $opts.IgnoreCase, $opts.Trim, $opts.IgnoreSpecialChars, $opts.IgnoreAllSpaces)
                        }
                        if (-not $isEqualOverall -and $diffPolicy -eq 'FuzzyContainment') {
                            $bmClean = ($bMerged -replace '[^\p{L}\p{Nd}\s]', ' ').Trim().ToLowerInvariant() -replace '\s+', ' '
                            $umClean = ($uMerged -replace '[^\p{L}\p{Nd}\s]', ' ').Trim().ToLowerInvariant() -replace '\s+', ' '
                            if (-not [string]::IsNullOrWhiteSpace($bmClean) -and -not [string]::IsNullOrWhiteSpace($umClean)) {
                                if ($bmClean.Contains($umClean) -or $umClean.Contains($bmClean)) {
                                    $isEqualOverall = $true
                                }
                            }
                        }
                        if ($isEqualOverall) {
                            continue
                        }
                    }

                    foreach ($baseCol in $bCols) {
                        if ($MetadataColumns -contains $baseCol) { continue }
                        $bVal = if ($matched.Values.ContainsKey($baseCol) -and $null -ne $matched.Values[$baseCol]) { $matched.Values[$baseCol].ToString() } else { '' }
                        $uVal = if ($projected.ContainsKey($baseCol) -and $null -ne $projected[$baseCol]) { $projected[$baseCol].ToString() } else { '' }

                        # Empty incoming cell semantics: keep base value, do not generate change
                        if ([string]::IsNullOrWhiteSpace($uVal) -and -not $emptyMeansClear) {
                            continue
                        }

                        $isEqual = [FastDiffHelper]::AreEqual($bVal, $uVal, $opts.IgnoreCase, $opts.Trim, $opts.IgnoreSpecialChars, $opts.IgnoreAllSpaces)
                        if (-not $isEqual -and $diffPolicy -eq 'NormalizePostalCode') {
                            $bNorm = ($bVal -replace '\b\d{2}-\d{3}\b|\b\d{5}\b', ' ') -replace '\s+', ' '
                            $uNorm = ($uVal -replace '\b\d{2}-\d{3}\b|\b\d{5}\b', ' ') -replace '\s+', ' '
                            $isEqual = [FastDiffHelper]::AreEqual($bNorm, $uNorm, $opts.IgnoreCase, $opts.Trim, $opts.IgnoreSpecialChars, $opts.IgnoreAllSpaces)
                        }
                        if (-not $isEqual -and $diffPolicy -eq 'FuzzyContainment') {
                            $bClean = ($bVal -replace '[^\p{L}\p{Nd}\s]', ' ').Trim().ToLowerInvariant() -replace '\s+', ' '
                            $uClean = ($uVal -replace '[^\p{L}\p{Nd}\s]', ' ').Trim().ToLowerInvariant() -replace '\s+', ' '
                            if (-not [string]::IsNullOrWhiteSpace($bClean) -and -not [string]::IsNullOrWhiteSpace($uClean)) {
                                if ($bClean.Contains($uClean) -or $uClean.Contains($bClean)) {
                                    $isEqual = $true
                                }
                            }
                        }

                        if (-not $isEqual) {
                            $colIdx = if ($BaseHeaders) { [System.Array]::IndexOf($BaseHeaders, $baseCol) } else { 0 }
                            $colIdx = if ($colIdx -lt 0) { 0 } else { $colIdx }
                            $colLetter = [FastExcelHelper]::ColIndexToName($colIdx)
                            $cellRef = "$colLetter$($matched.RowNumber)"
                            $changes.Add([ordered]@{
                                BaseColumn        = $baseCol
                                ColIndex          = $colIdx
                                CellRef           = $cellRef
                                OldValue          = $bVal
                                NewValue          = $uVal
                                OriginalNewValue  = $uVal
                                CustomEdited      = $false
                                SelectedForUpdate = $true # Q8: by default checked, can be deselected
                            })
                        }
                    }
                }

                $status = if ($changes.Count -gt 0) { 'Changed' } else { 'Unchanged' }
                $reviewItems.Add([PSCustomObject]@{
                    Index             = $idx
                    Status            = $status
                    Decision          = 'Pending'
                    AmbiguousReason   = $null
                    IncomingRow       = $inc
                    ProjectedRow      = $projected
                    MatchedBaseRow    = $matched
                    CandidateBaseRows = @()
                    Changes           = @($changes)
                    DecisionBy        = $env:USERNAME
                    DecisionUtc       = (Get-Date).ToUniversalTime().ToString('o')
                })
            }
        }
    }

    # 3. Detect removed/missing rows if enabled (Q3)
    if ($DetectRemoved -and -not [string]::IsNullOrEmpty($MarkDeletedColumn)) {
        foreach ($b in $normalizedBaseRows) {
            if (-not $claimedBaseRows.Contains($b.RowNumber)) {
                $curStatus = if ($b.Values.ContainsKey($MarkDeletedColumn) -and $null -ne $b.Values[$MarkDeletedColumn]) { $b.Values[$MarkDeletedColumn].ToString() } else { '' }
                if ($curStatus -ne $MarkDeletedValue) {
                    $idx++
                    $colIdx = if ($BaseHeaders) { [System.Array]::IndexOf($BaseHeaders, $MarkDeletedColumn) } else { -1 }
                    $colLetter = if ($colIdx -ge 0) { [FastExcelHelper]::ColIndexToName($colIdx) } else { 'A' }
                    $cellRef = "$colLetter$($b.RowNumber)"
                    $chg = [ordered]@{
                        BaseColumn        = $MarkDeletedColumn
                        ColIndex          = $colIdx
                        CellRef           = $cellRef
                        OldValue          = $curStatus
                        NewValue          = $MarkDeletedValue
                        SelectedForUpdate = $true
                    }
                    $reviewItems.Add([PSCustomObject]@{
                        Index             = $idx
                        Status            = 'Removed'
                        Decision          = 'Pending'
                        AmbiguousReason   = $null
                        IncomingRow       = $null
                        ProjectedRow      = @{ $MarkDeletedColumn = $MarkDeletedValue }
                        MatchedBaseRow    = $b
                        CandidateBaseRows = @()
                        Changes           = @($chg)
                        DecisionBy        = $env:USERNAME
                        DecisionUtc       = (Get-Date).ToUniversalTime().ToString('o')
                    })
                }
            }
        }
    }

    return $reviewItems
}

# ==============================================================================
# Region 6: Metadata Stamping Engine
# ==============================================================================
<#
.SYNOPSIS
    Resolves metadata token into formatted string value.

.DESCRIPTION
    Evaluates dynamic tokens (ChangeDate, SourceFileFullPath, SourceFileName, CurrentUser,
    SourceRowNumber, ImportBatchId) to stamp into base file audit columns.

.PARAMETER Token
    Token name to evaluate.

.PARAMETER Format
    Optional format string (e.g. for ChangeDate).

.PARAMETER IncomingRow
    Incoming row object containing source metadata.

.PARAMETER BatchId
    Session GUID.

.OUTPUTS
    System.String. Evaluated metadata string.
#>
function Get-StampedMetadataValue {
    param(
        [string]$Token,
        [string]$Format,
        [object]$IncomingRow,
        [string]$BatchId
    )
    switch ($Token) {
        'ChangeDate' {
            $fmt = if ([string]::IsNullOrWhiteSpace($Format)) { 'yyyy-MM-dd HH:mm' } else { $Format }
            return (Get-Date).ToString($fmt)
        }
        'SourceFileFullPath' {
            if ($IncomingRow -and $IncomingRow.SourceFilePath) { return $IncomingRow.SourceFilePath }
            return ''
        }
        'SourceFileName' {
            if ($IncomingRow -and $IncomingRow.SourceFilePath) { return (Split-Path $IncomingRow.SourceFilePath -Leaf) }
            return ''
        }
        'CurrentUser' {
            return $env:USERNAME
        }
        'SourceRowNumber' {
            if ($IncomingRow -and $IncomingRow.SourceRowNumber) { return $IncomingRow.SourceRowNumber.ToString() }
            return ''
        }
        'ImportBatchId' {
            return $BatchId
        }
        default {
            return ''
        }
    }
}

# ==============================================================================
# Region 7: Write-Back Engine (InPlace & SafeRewrite + Backup)
# ==============================================================================
<#
.SYNOPSIS
    Creates pre-write timestamped backup of the master base file.

.DESCRIPTION
    Copies the base file to the backup directory with timestamp format yyyy-MM-dd_HHmmss.bak.xlsx
    and prunes older backups according to retention count.

.PARAMETER BaseFilePath
    Path to the master base file.

.PARAMETER BackupDirectory
    Directory to store backups.

.PARAMETER RetentionCount
    Maximum number of backup files to keep.

.OUTPUTS
    System.String. Full path of created backup file.
#>
Set-Alias -Name 'New-BaseFileBackup' -Value 'Backup-BaseFile' -ErrorAction SilentlyContinue
function Backup-BaseFile {
    param(
        [string]$BaseFilePath,
        [string]$BackupDirectory,
        [int]$RetentionCount = 20,
        # P6: Auto-prune when total backup folder exceeds this threshold (MB). 0 = disabled.
        [int]$MaxBackupMb = 200
    )
    if (-not [string]::IsNullOrEmpty($BaseFilePath) -and -not (Test-Path $BaseFilePath)) { return $null }
    if ([string]::IsNullOrWhiteSpace($BackupDirectory)) {
        $baseDir = [System.IO.Path]::GetDirectoryName($BaseFilePath)
        $BackupDirectory = if (-not [string]::IsNullOrEmpty($baseDir)) { Join-Path $baseDir 'Backups' } else { '.\Backups' }
    }
    if (-not (Test-Path $BackupDirectory)) {
        [void][System.IO.Directory]::CreateDirectory($BackupDirectory)
    }

    $baseLeaf = [System.IO.Path]::GetFileNameWithoutExtension($BaseFilePath)
    $ext = [System.IO.Path]::GetExtension($BaseFilePath)
    $ts = (Get-Date).ToString('yyyy-MM-dd_HHmmss')
    $backupFileName = "$baseLeaf.$ts.bak$ext"
    $backupPath = Join-Path $BackupDirectory $backupFileName

    Copy-Item -Path $BaseFilePath -Destination $backupPath -Force

    # Enforce retention count (by file count)
    $filter = "$baseLeaf.*.bak$ext"
    $backups = Get-ChildItem -Path $BackupDirectory -Filter $filter -File | Sort-Object CreationTime
    if ($backups.Count -gt $RetentionCount) {
        $toDeleteCount = $backups.Count - $RetentionCount
        for ($i = 0; $i -lt $toDeleteCount; $i++) {
            try { Remove-Item -Force $backups[$i].FullName } catch { }
        }
        # Refresh list after count-pruning
        $backups = Get-ChildItem -Path $BackupDirectory -Filter $filter -File | Sort-Object CreationTime
    }

    # P6: Auto-prune oldest backups when total folder size exceeds MaxBackupMb
    $sizeWarning = $null
    if ($MaxBackupMb -gt 0) {
        $allBackups = @(Get-ChildItem -Path $BackupDirectory -Filter '*.bak.*' -File | Sort-Object CreationTime)
        $totalMb = [math]::Round(($allBackups | Measure-Object -Property Length -Sum).Sum / 1MB, 1)
        if ($totalMb -gt $MaxBackupMb) {
            # Prune oldest until under threshold
            $idx = 0
            while ($totalMb -gt $MaxBackupMb -and $idx -lt $allBackups.Count) {
                $victim = $allBackups[$idx]
                $sizeMb = [math]::Round($victim.Length / 1MB, 1)
                try {
                    Remove-Item -Force $victim.FullName
                    $totalMb -= $sizeMb
                } catch { }
                $idx++
            }
            $sizeWarning = "Folder kopii zapasowych przekroczyl ${MaxBackupMb} MB. Najstarsze kopie zostaly automatycznie usuniete. Rozmiar po czyszczeniu: ${totalMb} MB."
        }
    }

    return [PSCustomObject]@{
        BackupPath   = $backupPath
        SizeWarning  = $sizeWarning
    }
}

<#
.SYNOPSIS
    Writes accepted review items to the master base file.

.DESCRIPTION
    Performs pre-write backup, ensures metadata columns exist in header row, and applies
    changes using InPlace OpenXML stream patching or SafeRewrite.

.PARAMETER BaseFilePath
    Path to master base file.

.PARAMETER BaseSheet
    Name of worksheet to update.

.PARAMETER ReviewItems
    List of accepted review items.

.PARAMETER AppConfig
    Application configuration hashtable.

.PARAMETER BaseHeaders
    Array of base file headers.

.PARAMETER BatchId
    Unique import batch GUID.

.OUTPUTS
    Hashtable with Success, BackupPath, UpdatedCells, AddedRows, and BatchId.
#>
function Invoke-MasterWriteBack {
    param(
        [string]$BaseFilePath,
        [string]$BaseSheet,
        [Alias('AcceptedItems')]
        [System.Collections.IList]$ReviewItems,     # Only Accepted items
        [hashtable]$AppConfig,
        [string[]]$BaseHeaders,
        [string]$BatchId = ([guid]::NewGuid().ToString())
    )

    if (-not (Test-Path $BaseFilePath)) { throw "Base file not found: $BaseFilePath" }

    # 1. Backup first
    $backupResult = Backup-BaseFile -BaseFilePath $BaseFilePath -BackupDirectory $AppConfig.BackupDirectory -RetentionCount $AppConfig.BackupRetentionCount -MaxBackupMb ($(if ($AppConfig.MaxBackupMb) { $AppConfig.MaxBackupMb } else { 200 }))
    $backupPath = if ($backupResult) { $backupResult.BackupPath } else { $null }
    $backupSizeWarning = if ($backupResult) { $backupResult.SizeWarning } else { $null }

    if (-not $BaseHeaders -or $BaseHeaders.Length -eq 0) {
        $BaseHeaders = [FastExcelHelper]::GetHeaders($BaseFilePath, $BaseSheet)
    }

    $metaCols = @($AppConfig.MetadataColumns)
    $headersList = [System.Collections.Generic.List[string]]::new([string[]]$BaseHeaders)

    $getMetaProp = {
        param($mObj, [string]$prop)
        if ($null -eq $mObj) { return $null }
        if ($mObj -is [System.Collections.IDictionary]) {
            if ($mObj.Contains($prop)) { return $mObj[$prop] }
            return $null
        }
        if ($mObj.PSObject.Properties[$prop]) {
            return $mObj.$prop
        }
        return $null
    }

    # 2. Check for missing metadata columns in base headers
    $missingMeta = [System.Collections.Generic.List[object]]::new()
    foreach ($m in $metaCols) {
        $colName = & $getMetaProp $m 'BaseColumn'
        if (-not [string]::IsNullOrWhiteSpace($colName) -and -not $headersList.Contains($colName)) {
            $missingMeta.Add($m)
            $headersList.Add($colName)
        }
    }

    $isCsv = $BaseFilePath -match '\.csv$'
    $writeMode = if ($isCsv) { 'SafeRewrite' } elseif ($AppConfig -and $AppConfig.WriteMode) { $AppConfig.WriteMode } else { 'InPlace' }

    $acceptedItems = [System.Collections.Generic.List[object]]::new()
    foreach ($it in $ReviewItems) {
        $actualItem = if ($it.PSObject.Properties['Record'] -and $it.Record) { $it.Record } else { $it }
        $dec = if ($it.PSObject.Properties['Decision'] -and $it.Decision) {
            $it.Decision
        } elseif ($actualItem.PSObject.Properties['Decision'] -and $actualItem.Decision) {
            $actualItem.Decision
        } else {
            $actualItem.Status
        }
        if ($dec -eq 'Accepted' -or $it.Status -eq 'Accepted' -or $actualItem.Status -eq 'Accepted') {
            $acceptedItems.Add($actualItem)
        }
    }
    if ($acceptedItems.Count -eq 0 -and $ReviewItems.Count -gt 0) {
        foreach ($it in $ReviewItems) {
            $actualItem = if ($it.PSObject.Properties['Record'] -and $it.Record) { $it.Record } else { $it }
            $acceptedItems.Add($actualItem)
        }
    }
    if ($acceptedItems.Count -eq 0) {
        return @{
            Success      = $true
            BackupPath   = $backupPath
            UpdatedCells = 0
            AddedRows    = 0
            BatchId      = $BatchId
        }
    }

    if ($writeMode -eq 'InPlace') {
        $ops = New-Object 'System.Collections.Generic.List[RowOp]'

        # Add missing headers to row 1 if any
        if ($missingMeta.Count -gt 0) {
            $headerPatch = New-Object RowOp
            $headerPatch.Type = 'PatchCell'
            $headerPatch.RowNumber = 1
            foreach ($m in $missingMeta) {
                $cIdx = $headersList.IndexOf($m.BaseColumn)
                $headerPatch.Cells.Add($cIdx, $m.BaseColumn)
            }
            $ops.Add($headerPatch)
        }

        # Determine last row number in base file
        $baseRaw = [FastExcelHelper]::ReadSheet($BaseFilePath, $BaseSheet)
        $lastRowNum = 1
        foreach ($r in $baseRaw) {
            if ([int]$r._RowNumber -gt $lastRowNum) { $lastRowNum = [int]$r._RowNumber }
        }

        $updatedCellsCount = 0
        $addedRowsCount = 0

        foreach ($item in $acceptedItems) {
            if ($item.Status -eq 'Unchanged' -and (-not $item.Changes -or $item.Changes.Count -eq 0)) {
                continue
            }
            if ($item.MatchedBaseRow) {
                # Patch existing row
                $patchOp = New-Object RowOp
                $patchOp.Type = 'PatchCell'
                $patchOp.RowNumber = $item.MatchedBaseRow.RowNumber

                # Apply selected changes (Q8)
                foreach ($chg in $item.Changes) {
                    if ($chg.SelectedForUpdate -ne $false) {
                        $cIdx = $headersList.IndexOf($chg.BaseColumn)
                        if ($cIdx -ge 0) {
                            $patchOp.Cells[$cIdx] = if ($null -ne $chg.NewValue) { $chg.NewValue.ToString() } else { '' }
                            $updatedCellsCount++
                        }
                    }
                }

                # Stamp metadata
                foreach ($m in $metaCols) {
                    $mCol = & $getMetaProp $m 'BaseColumn'
                    if ([string]::IsNullOrWhiteSpace($mCol)) { continue }
                    $cIdx = $headersList.IndexOf($mCol)
                    if ($cIdx -ge 0) {
                        $mTok = & $getMetaProp $m 'Token'
                        $mFmt = & $getMetaProp $m 'Format'
                        $metaVal = Get-StampedMetadataValue -Token $mTok -Format $mFmt -IncomingRow $item.IncomingRow -BatchId $BatchId
                        $patchOp.Cells[$cIdx] = $metaVal
                    }
                }

                $ops.Add($patchOp)
            } else {
                # Append new row
                $lastRowNum++
                $appendOp = New-Object RowOp
                $appendOp.Type = 'AppendRow'
                $appendOp.RowNumber = $lastRowNum

                # Mapped projected values
                foreach ($col in $headersList) {
                    $cIdx = $headersList.IndexOf($col)
                    $val = if ($item.ProjectedRow.ContainsKey($col)) { $item.ProjectedRow[$col] } else { '' }
                    if (-not [string]::IsNullOrEmpty($val)) {
                        $appendOp.Cells[$cIdx] = $val.ToString()
                    }
                }

                # Stamp metadata
                foreach ($m in $metaCols) {
                    $mCol = & $getMetaProp $m 'BaseColumn'
                    if ([string]::IsNullOrWhiteSpace($mCol)) { continue }
                    $cIdx = $headersList.IndexOf($mCol)
                    if ($cIdx -ge 0) {
                        $mTok = & $getMetaProp $m 'Token'
                        $mFmt = & $getMetaProp $m 'Format'
                        $metaVal = Get-StampedMetadataValue -Token $mTok -Format $mFmt -IncomingRow $item.IncomingRow -BatchId $BatchId
                        $appendOp.Cells[$cIdx] = $metaVal
                    }
                }

                $ops.Add($appendOp)
                $addedRowsCount++
            }
        }

        # Execute InPlace write
        # P5: Catch the INPLACE_GUARD_WARN signal and fall back to SafeRewrite automatically
        try {
            [EditExcelHelper]::WriteChanges($BaseFilePath, $BaseSheet, $ops)
        } catch [System.InvalidOperationException] {
            if ($_.Exception.Message -like 'INPLACE_GUARD_WARN:*') {
                # Fall back to SafeRewrite
                $writeMode = 'SafeRewrite'
                Write-Warning "InPlace guard triggered. Falling back to SafeRewrite mode: $($_.Exception.Message)"
            } else {
                throw
            }
        }

        if ($writeMode -ne 'SafeRewrite') {
            return @{
                Success          = $true
                BackupPath       = $backupPath
                UpdatedCells     = $updatedCellsCount
                AddedRows        = $addedRowsCount
                BatchId          = $BatchId
                BackupSizeWarning = $backupSizeWarning
            }
        }
        # SafeRewrite Mode (or CSV)
        $baseRaw = [FastExcelHelper]::ReadSheet($BaseFilePath, $BaseSheet)
        $rowLookup = @{}
        foreach ($r in $baseRaw) {
            $rowLookup[[int]$r._RowNumber] = $r
        }

        $updatedCellsCount = 0
        $addedRowsCount = 0

        # Apply changes to in-memory objects
        foreach ($item in $acceptedItems) {
            if ($item.Status -eq 'Unchanged' -and (-not $item.Changes -or $item.Changes.Count -eq 0)) {
                continue
            }
            if ($item.MatchedBaseRow) {
                $targetR = $rowLookup[$item.MatchedBaseRow.RowNumber]
                if ($targetR) {
                    foreach ($chg in $item.Changes) {
                        if ($chg.SelectedForUpdate -ne $false) {
                            $col = $chg.BaseColumn
                            if (-not $targetR.PSObject.Properties[$col]) {
                                $targetR.PSObject.Properties.Add([System.Management.Automation.PSNoteProperty]::new($col, $chg.NewValue))
                            } else {
                                $targetR.$col = $chg.NewValue
                            }
                            $updatedCellsCount++
                        }
                    }
                    # Stamp metadata
                    foreach ($m in $metaCols) {
                        $col = & $getMetaProp $m 'BaseColumn'
                        if ([string]::IsNullOrWhiteSpace($col)) { continue }
                        $mTok = & $getMetaProp $m 'Token'
                        $mFmt = & $getMetaProp $m 'Format'
                        $metaVal = Get-StampedMetadataValue -Token $mTok -Format $mFmt -IncomingRow $item.IncomingRow -BatchId $BatchId
                        if (-not $targetR.PSObject.Properties[$col]) {
                            $targetR.PSObject.Properties.Add([System.Management.Automation.PSNoteProperty]::new($col, $metaVal))
                        } else {
                            $targetR.$col = $metaVal
                        }
                    }
                }
            } else {
                # New row
                $newObj = New-Object PSObject
                foreach ($col in $headersList) {
                    $val = if ($item.ProjectedRow.ContainsKey($col)) { $item.ProjectedRow[$col] } else { '' }
                    $newObj.PSObject.Properties.Add([System.Management.Automation.PSNoteProperty]::new($col, $val))
                }
                # Stamp metadata
                foreach ($m in $metaCols) {
                    $col = & $getMetaProp $m 'BaseColumn'
                    if ([string]::IsNullOrWhiteSpace($col)) { continue }
                    $mTok = & $getMetaProp $m 'Token'
                    $mFmt = & $getMetaProp $m 'Format'
                    $metaVal = Get-StampedMetadataValue -Token $mTok -Format $mFmt -IncomingRow $item.IncomingRow -BatchId $BatchId
                    if ($newObj.PSObject.Properties[$col]) {
                        $newObj.$col = $metaVal
                    } else {
                        $newObj.PSObject.Properties.Add([System.Management.Automation.PSNoteProperty]::new($col, $metaVal))
                    }
                }
                [void]$baseRaw.Add($newObj)
                $addedRowsCount++
            }
        }

        # Export (enforce UTF-8 with BOM for Excel compatibility)
        if ($isCsv) {
            $csvLines = $baseRaw | Select-Object -Property $headersList | ConvertTo-Csv -NoTypeInformation -Delimiter ';'
            [System.IO.File]::WriteAllLines($BaseFilePath, $csvLines, [System.Text.UTF8Encoding]::new($true))
        } else {
            [FastExcelHelper]::ExportToExcel($BaseFilePath, $baseRaw, $true, $true, $true)
        }

        return @{
            Success           = $true
            BackupPath        = $backupPath
            UpdatedCells      = $updatedCellsCount
            AddedRows         = $addedRowsCount
            BatchId           = $BatchId
            BackupSizeWarning = $backupSizeWarning
        }
    }
}

# ==============================================================================
# Region 8: Logging Engine (Write-ImportLog)
# ==============================================================================
<#
.SYNOPSIS
    Generates audit logs for the import session.

.DESCRIPTION
    Writes session JSONL event log, human-readable summary TXT report, and appends to
    the global masterlog.jsonl audit trail with optional PII redaction.

.PARAMETER LogDirectory
    Target directory for log storage.

.PARAMETER BatchId
    Session GUID.

.PARAMETER ReviewItems
    Collection of all review items and decisions.

.PARAMETER WriteBackResult
    Write-back execution summary.

.PARAMETER BaseFilePath
    Path of updated base file.

.PARAMETER RedactNames
    Whether to mask personal names and addresses in logs (GDPR/PII).

.OUTPUTS
    Hashtable containing JsonlPath, TxtPath, and MasterLogPath.
#>
function Write-ImportLog {
    param(
        [string]$LogDirectory,
        [string]$BatchId,
        [System.Collections.IList]$ReviewItems,
        [hashtable]$WriteBackResult,
        [string]$BaseFilePath,
        [bool]$RedactNames = $false,
        [bool]$LogChangesToBaseSheet = $false,
        [string]$BaseSheetLogName = 'ImportLog',
        [hashtable]$AppConfig = $null
    )

    if ($AppConfig) {
        if ($null -ne $AppConfig.LogChangesToBaseSheet) { $LogChangesToBaseSheet = [bool]$AppConfig.LogChangesToBaseSheet }
        if (-not [string]::IsNullOrWhiteSpace($AppConfig.BaseSheetLogName)) { $BaseSheetLogName = $AppConfig.BaseSheetLogName }
    }

    if (-not (Test-Path $LogDirectory)) {
        [void][System.IO.Directory]::CreateDirectory($LogDirectory)
    }

    $tsStr = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $nowStr = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $jsonlPath = Join-Path $LogDirectory "import_${tsStr}_$BatchId.jsonl"
    $txtPath   = Join-Path $LogDirectory "import_${tsStr}_$BatchId.txt"
    $masterJsonlPath = Join-Path $LogDirectory "masterlog.jsonl"

    $jsonlLines = [System.Collections.Generic.List[string]]::new()
    $txtLines   = [System.Collections.Generic.List[string]]::new()
    $sheetLogRows = [System.Collections.Generic.List[object]]::new()

    $txtLines.Add("================================================================================")
    $txtLines.Add("EXCEL MASTER UPDATER - IMPORT SESSION LOG")
    $txtLines.Add("Session Date/Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | User: $env:USERNAME")
    $txtLines.Add("Batch ID: $BatchId")
    $txtLines.Add("Base File: $BaseFilePath")
    $txtLines.Add("Backup File: $($WriteBackResult.BackupPath)")
    $txtLines.Add("================================================================================")
    $txtLines.Add("")

    $cNew = 0; $cChg = 0; $cAcc = 0; $cRej = 0; $cSkip = 0; $cAmb = 0; $cUnch = 0

    $piiPattern = '(?i)(nazwisko|imie|imię|adres|pesel|telefon|mail|email|name|surname|firstname|lastname|address|ssn|phone|street|nachname|vorname|anschrift|geburtsdatum)'

    foreach ($item in $ReviewItems) {
        $rec = if ($item.PSObject.Properties['Record'] -and $item.Record) { $item.Record } else { $item }
        $st = if ($rec.Status) { $rec.Status } else { $item.Status }
        $dec = if ($item.PSObject.Properties['Decision'] -and $item.Decision) {
            $item.Decision
        } elseif ($rec.PSObject.Properties['Decision'] -and $rec.Decision) {
            $rec.Decision
        } else { $st }

        switch ($st) {
            'New'        { $cNew++ }
            'Changed'    { $cChg++ }
            'Accepted'   { $cAcc++ }
            'Rejected'   { $cRej++ }
            'Skipped'    { $cSkip++ }
            'Ambiguous'  { $cAmb++ }
            'Unchanged'  { $cUnch++ }
        }

        $isAccepted = ($st -eq 'Accepted' -or $dec -eq 'Accepted' -or $dec -eq 'Accept')
        $keyVal = if ($item.PSObject.Properties['Title'] -and -not [string]::IsNullOrWhiteSpace($item.Title)) {
            $item.Title
        } elseif ($rec.MatchedBaseRow -and $rec.MatchedBaseRow.Values) {
            ($rec.MatchedBaseRow.Values.Values | Select-Object -First 1)
        } elseif ($rec.IncomingRow -and $rec.IncomingRow.Raw) {
            ($rec.IncomingRow.Raw.PSObject.Properties | Select-Object -ExpandProperty Value -First 1)
        } else { '' }

        if ($isAccepted) {
            if ($rec.MatchedBaseRow) {
                $baseRowNum = $rec.MatchedBaseRow.RowNumber.ToString()
                $srcFile = if ($rec.IncomingRow -and $rec.IncomingRow.SourceFilePath) { Split-Path $rec.IncomingRow.SourceFilePath -Leaf } else { '' }
                $srcRowNum = if ($rec.IncomingRow -and $rec.IncomingRow.SourceRowNumber) { $rec.IncomingRow.SourceRowNumber.ToString() } else { '' }
                foreach ($chg in $rec.Changes) {
                    if ($chg.SelectedForUpdate -ne $false) {
                        $oldV = if ($RedactNames -and ($chg.BaseColumn -match $piiPattern)) { '***' } else { $chg.OldValue }
                        $newV = if ($RedactNames -and ($chg.BaseColumn -match $piiPattern)) { '***' } else { $chg.NewValue }
                        $evt = [ordered]@{
                            ts         = (Get-Date).ToUniversalTime().ToString('o')
                            user       = $env:USERNAME
                            batch      = $BatchId
                            action     = 'CellChanged'
                            baseRow    = $rec.MatchedBaseRow.RowNumber
                            cell       = $chg.CellRef
                            column     = $chg.BaseColumn
                            old        = $oldV
                            new        = $newV
                            sourceFile = $rec.IncomingRow.SourceFilePath
                            sourceRow  = $rec.IncomingRow.SourceRowNumber
                        }
                        $jsonlLines.Add(($evt | ConvertTo-Json -Compress))
                        $txtLines.Add("[CHANGE] Base row #$($rec.MatchedBaseRow.RowNumber) Cell $($chg.CellRef) [$($chg.BaseColumn)]: '$oldV' -> '$newV' (Source row: $($rec.IncomingRow.SourceRowNumber))")
                        $sheetLogRows.Add(@($nowStr, $BatchId, $env:USERNAME, 'CellChanged', $baseRowNum, $keyVal, $chg.BaseColumn, $oldV, $newV, $srcFile, $srcRowNum))
                    }
                }
            } elseif ($st -eq 'Removed') {
                $baseRowNum = if ($rec.MatchedBaseRow) { $rec.MatchedBaseRow.RowNumber.ToString() } else { '' }
                $evt = [ordered]@{
                    ts         = (Get-Date).ToUniversalTime().ToString('o')
                    user       = $env:USERNAME
                    batch      = $BatchId
                    action     = 'Removed'
                    baseRow    = $baseRowNum
                }
                $jsonlLines.Add(($evt | ConvertTo-Json -Compress))
                $txtLines.Add("[REMOVED] Record base row #$baseRowNum marked as removed")
                $sheetLogRows.Add(@($nowStr, $BatchId, $env:USERNAME, 'Removed', $baseRowNum, $keyVal, '(Status)', '', 'Removed', '', ''))
            } else {
                $srcFile = if ($rec.IncomingRow -and $rec.IncomingRow.SourceFilePath) { Split-Path $rec.IncomingRow.SourceFilePath -Leaf } else { '' }
                $srcRowNum = if ($rec.IncomingRow -and $rec.IncomingRow.SourceRowNumber) { $rec.IncomingRow.SourceRowNumber.ToString() } else { '' }
                $evt = [ordered]@{
                    ts         = (Get-Date).ToUniversalTime().ToString('o')
                    user       = $env:USERNAME
                    batch      = $BatchId
                    action     = 'AddedRow'
                    baseRow    = 'New'
                    sourceFile = $rec.IncomingRow.SourceFilePath
                    sourceRow  = $rec.IncomingRow.SourceRowNumber
                }
                $jsonlLines.Add(($evt | ConvertTo-Json -Compress))
                $txtLines.Add("[NEW ROW] Added new record from source row #$($rec.IncomingRow.SourceRowNumber)")
                $newSummary = if ($rec.ProjectedRow) {
                    ($rec.ProjectedRow.Keys | ForEach-Object {
                        $v = $rec.ProjectedRow[$_]
                        if ($RedactNames -and ($_ -match $piiPattern)) { $v = '***' }
                        "$($_): $v"
                    }) -join '; '
                } else { '' }
                $sheetLogRows.Add(@($nowStr, $BatchId, $env:USERNAME, 'AddedRow', 'New', $keyVal, '(All)', '', $newSummary, $srcFile, $srcRowNum))
            }
        } elseif ($st -eq 'Rejected' -or $dec -eq 'Rejected') {
            $srcRowNum = if ($rec.IncomingRow -and $rec.IncomingRow.SourceRowNumber) { $rec.IncomingRow.SourceRowNumber } else { '' }
            $evt = [ordered]@{
                ts         = (Get-Date).ToUniversalTime().ToString('o')
                user       = $env:USERNAME
                batch      = $BatchId
                action     = 'Rejected'
                sourceRow  = $srcRowNum
                reason     = 'Rejected by user'
            }
            $jsonlLines.Add(($evt | ConvertTo-Json -Compress))
            $txtLines.Add("[REJECTED] Source row #$srcRowNum")
        } elseif ($st -eq 'Skipped' -or $dec -eq 'Skipped') {
            $evt = [ordered]@{
                ts         = (Get-Date).ToUniversalTime().ToString('o')
                user       = $env:USERNAME
                batch      = $BatchId
                action     = 'Skipped'
                sourceRow  = $item.IncomingRow.SourceRowNumber
            }
            $jsonlLines.Add(($evt | ConvertTo-Json -Compress))
        }
    }

    # Summary event
    $summaryEvt = [ordered]@{
        ts           = (Get-Date).ToUniversalTime().ToString('o')
        action       = 'SessionSummary'
        incoming     = $ReviewItems.Count
        new          = $cNew
        changed      = $cChg
        accepted     = $cAcc
        rejected     = $cRej
        skipped      = $cSkip
        ambiguous    = $cAmb
        unchanged    = $cUnch
        baseFile     = $BaseFilePath
        backupFile   = $WriteBackResult.BackupPath
        batch        = $BatchId
    }
    $jsonlLines.Add(($summaryEvt | ConvertTo-Json -Compress))

    $txtLines.Add("")
    $txtLines.Add("--------------------------------------------------------------------------------")
    $txtLines.Add("SESSION SUMMARY:")
    $txtLines.Add("Total rows in incoming file: $($ReviewItems.Count)")
    $txtLines.Add("Accepted: $cAcc (Added rows: $($WriteBackResult.AddedRows), Updated cells: $($WriteBackResult.UpdatedCells))")
    $txtLines.Add("Rejected: $cRej | Skipped: $cSkip | Ambiguous: $cAmb | Unchanged: $cUnch")
    $txtLines.Add("--------------------------------------------------------------------------------")

    [System.IO.File]::WriteAllLines($jsonlPath, $jsonlLines, [System.Text.UTF8Encoding]::new($true))
    [System.IO.File]::WriteAllLines($txtPath, $txtLines, [System.Text.UTF8Encoding]::new($true))

    # Append to master log
    $sw = [System.IO.File]::AppendText($masterJsonlPath)
    try {
        foreach ($line in $jsonlLines) {
            $sw.WriteLine($line)
        }
    } finally {
        $sw.Dispose()
    }

    $baseSheetWritten = $null
    if ($LogChangesToBaseSheet -and $BaseFilePath -and (Test-Path $BaseFilePath)) {
        if ($BaseFilePath -match '\.csv$') {
            Write-Warning "Base file is in CSV format. Sheet logging to '$BaseSheetLogName' is only supported for Excel (.xlsx) files."
        } elseif ($sheetLogRows.Count -gt 0) {
            $sheetName = if (-not [string]::IsNullOrWhiteSpace($BaseSheetLogName)) { $BaseSheetLogName } else { 'ImportLog' }
            $sheetName = [regex]::Replace($sheetName, '[\\/\?\*:[\]]', '_')
            if ($sheetName.Length -gt 31) { $sheetName = $sheetName.Substring(0, 31) }
            $logHeaders = @("Timestamp", "BatchId", "User", "Action", "BaseRow", "Key", "Column", "OldValue", "NewValue", "SourceFile", "SourceRow")
            [EditExcelHelper]::AppendWorksheetLog($BaseFilePath, $sheetName, $logHeaders, $sheetLogRows)
            $baseSheetWritten = $sheetName
        }
    }

    return @{
        JsonlPath    = $jsonlPath
        TxtPath      = $txtPath
        BaseSheetLog = $baseSheetWritten
    }
}

<#
.SYNOPSIS
    Exports review and discrepancy results to an external report file (.html, .xlsx, or .csv).
#>
function Export-ReviewReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IList]$ReviewItems,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,

        [string]$BaseFilePath = '',
        [string]$IncomingPath = '',
        [string]$BatchId = ''
    )

    if (-not $ReviewItems -or $ReviewItems.Count -eq 0) {
        throw "No review items to export."
    }

    $destDir = Split-Path -Parent $DestinationPath
    if ($destDir -and -not (Test-Path $destDir)) {
        [void][System.IO.Directory]::CreateDirectory($destDir)
    }

    if (-not $script:LanguagesCatalog -or $script:LanguagesCatalog.Count -eq 0) {
        if (Get-Command Import-LanguageCatalog -ErrorAction SilentlyContinue) {
            Import-LanguageCatalog
        }
    }

    $cIndex     = Get-UiString 'ReportColIndex' -Default 'Index'
    $cStatus    = Get-UiString 'ReportColStatus' -Default 'Status'
    $cDecision  = Get-UiString 'ReportColDecision' -Default 'Decision'
    $cId        = Get-UiString 'ReportColIdentifier' -Default 'Identifier'
    $cBaseRow   = Get-UiString 'ReportColBaseRow' -Default 'Base Row'
    $cSourceRow = Get-UiString 'ReportColSourceRow' -Default 'Source Row'
    $cChgCount  = Get-UiString 'ReportColChangeCount' -Default 'Change Count'
    $cChgSumm   = Get-UiString 'ReportColChangeSummary' -Default 'Diff Summary'
    $cAmbReason = Get-UiString 'ReportColAmbiguousReason' -Default 'Ambiguity Reason'

    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($it in $ReviewItems) {
        $rec = if ($it.PSObject.Properties['Record']) { $it.Record } else { $it }
        $dec = if ($it.PSObject.Properties['Decision']) { $it.Decision } else { $rec.Status }
        $idx = if ($it.PSObject.Properties['Index']) { $it.Index } elseif ($it.PSObject.Properties['IndexStr']) { $it.IndexStr } else { '' }
        $title = if ($it.PSObject.Properties['Title']) { $it.Title } else { '' }

        $chgSumm = if ($rec.Changes -and $rec.Changes.Count -gt 0) {
            ($rec.Changes | ForEach-Object { "$($_.BaseColumn): '$($_.OldValue)' -> '$($_.NewValue)'" }) -join '; '
        } else { '' }

        $rowObj = [ordered]@{}
        $rowObj[$cIndex]     = $idx
        $rowObj[$cStatus]    = $rec.Status
        $rowObj[$cDecision]  = $dec
        $rowObj[$cId]        = $title
        $rowObj[$cBaseRow]   = if ($rec.MatchedBaseRow) { $rec.MatchedBaseRow.RowNumber } else { '' }
        $rowObj[$cSourceRow] = if ($rec.IncomingRow) { $rec.IncomingRow.SourceRowNumber } else { '' }
        $rowObj[$cChgCount]  = if ($rec.Changes) { $rec.Changes.Count } else { 0 }
        $rowObj[$cChgSumm]   = $chgSumm
        $rowObj[$cAmbReason] = if ($rec.AmbiguousReason) { $rec.AmbiguousReason } else { '' }
        $rows.Add([PSCustomObject]$rowObj)
    }

    if ($DestinationPath -match '\.html?$') {
        # Standalone responsive HTML Diff Report (Opportunity O4)
        $html = [System.Text.StringBuilder]::new()
        [void]$html.AppendLine('<!DOCTYPE html>')
        $currLang = if ($script:CurrentLanguage) { $script:CurrentLanguage } else { 'pl' }
        [void]$html.AppendLine("<html lang=""$currLang"">")
        [void]$html.AppendLine('<head>')
        [void]$html.AppendLine('  <meta charset="UTF-8">')
        [void]$html.AppendLine('  <meta name="viewport" content="width=device-width, initial-scale=1.0">')
        $reportTitle = Get-UiString 'ReportHtmlTitle' -Default 'Excel Master Updater &mdash; Review & Diff Report'
        [void]$html.AppendLine("  <title>$reportTitle</title>")
        [void]$html.AppendLine('  <style>')
        [void]$html.AppendLine('    :root { --bg: #0F172A; --card: #1E293B; --border: #334155; --text: #F8FAFC; --text-muted: #94A3B8; --green: #10B981; --amber: #F59E0B; --red: #EF4444; --blue: #3B82F6; }')
        [void]$html.AppendLine('    body { margin: 0; padding: 24px; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; background: var(--bg); color: var(--text); line-height: 1.5; }')
        [void]$html.AppendLine('    .container { max-width: 1300px; margin: 0 auto; }')
        [void]$html.AppendLine('    header { padding-bottom: 20px; border-bottom: 1px solid var(--border); margin-bottom: 24px; }')
        [void]$html.AppendLine('    h1 { margin: 0 0 8px 0; font-size: 24px; font-weight: 700; color: #FFFFFF; }')
        [void]$html.AppendLine('    .meta { font-size: 13px; color: var(--text-muted); }')
        [void]$html.AppendLine('    .meta span { margin-right: 20px; display: inline-block; margin-bottom: 4px; }')
        [void]$html.AppendLine('    .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 14px; margin-bottom: 28px; }')
        [void]$html.AppendLine('    .kpi-card { background: var(--card); border: 1px solid var(--border); border-radius: 8px; padding: 14px 18px; }')
        [void]$html.AppendLine('    .kpi-card .num { font-size: 28px; font-weight: 800; line-height: 1.1; margin-bottom: 4px; }')
        [void]$html.AppendLine('    .kpi-card .lbl { font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; color: var(--text-muted); }')
        [void]$html.AppendLine('    .kpi-new .num { color: #34D399; }')
        [void]$html.AppendLine('    .kpi-chg .num { color: #FBBF24; }')
        [void]$html.AppendLine('    .kpi-amb .num { color: #F87171; }')
        [void]$html.AppendLine('    .kpi-acc .num { color: #60A5FA; }')
        [void]$html.AppendLine('    .section-title { font-size: 18px; font-weight: 600; margin: 24px 0 12px 0; }')
        [void]$html.AppendLine('    table { width: 100%; border-collapse: collapse; background: var(--card); border: 1px solid var(--border); border-radius: 8px; overflow: hidden; font-size: 13px; }')
        [void]$html.AppendLine('    th { background: #162032; text-align: left; padding: 10px 14px; font-size: 12px; text-transform: uppercase; color: var(--text-muted); border-bottom: 1px solid var(--border); }')
        [void]$html.AppendLine('    td { padding: 10px 14px; border-bottom: 1px solid #233146; vertical-align: top; }')
        [void]$html.AppendLine('    tr:last-child td { border-bottom: none; }')
        [void]$html.AppendLine('    .badge { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 700; text-transform: uppercase; }')
        [void]$html.AppendLine('    .b-new { background: #064E3B; color: #A7F3D0; }')
        [void]$html.AppendLine('    .b-chg { background: #713F12; color: #FDE68A; }')
        [void]$html.AppendLine('    .b-amb { background: #7F1D1D; color: #FECACA; }')
        [void]$html.AppendLine('    .b-rem { background: #881337; color: #FECDD3; }')
        [void]$html.AppendLine('    .b-unch { background: #1E293B; color: #94A3B8; border: 1px solid #334155; }')
        [void]$html.AppendLine('    footer { margin-top: 40px; padding-top: 16px; border-top: 1px solid var(--border); font-size: 12px; color: var(--text-muted); text-align: center; }')
        [void]$html.AppendLine('    @media print { body { background: #FFF; color: #000; } .kpi-card, table { border-color: #CCC; background: #FFF; color: #000; } th { background: #EEE; color: #000; } }')
        [void]$html.AppendLine('  </style>')
        [void]$html.AppendLine('</head>')
        [void]$html.AppendLine('<body>')
        [void]$html.AppendLine('  <div class="container">')
        [void]$html.AppendLine('    <header>')
        [void]$html.AppendLine("      <h1>$reportTitle</h1>")
        $nowStr = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        $bFileEnc = [System.Net.WebUtility]::HtmlEncode($BaseFilePath)
        $iFileEnc = [System.Net.WebUtility]::HtmlEncode($IncomingPath)
        $lblDate  = Get-UiString 'ReportMetaDate' -Default 'Date:'
        $lblBase  = Get-UiString 'ReportMetaBase' -Default 'Base File:'
        $lblInc   = Get-UiString 'ReportMetaIncoming' -Default 'Changes File:'
        $lblBatch = Get-UiString 'ReportMetaBatch' -Default 'Batch ID:'
        [void]$html.AppendLine("      <div class=""meta""><span><b>$lblDate</b> $nowStr</span><span><b>$lblBase</b> $bFileEnc</span><span><b>$lblInc</b> $iFileEnc</span>$(if ($BatchId) { "<span><b>$lblBatch</b> $BatchId</span>" } else { '' })</div>")
        [void]$html.AppendLine('    </header>')

        $kNew = @($rows | Where-Object { $_.$cStatus -eq 'New' }).Count
        $kChg = @($rows | Where-Object { $_.$cStatus -eq 'Changed' }).Count
        $kAmb = @($rows | Where-Object { $_.$cStatus -eq 'Ambiguous' }).Count
        $kAcc = @($rows | Where-Object { $_.$cDecision -eq 'Accepted' }).Count

        $lblAllRows = Get-UiString 'ReportKpiAllRows' -Default 'All Rows'
        $lblNewRows = Get-UiString 'ReportKpiNewRecords' -Default 'New Records'
        $lblChgRows = Get-UiString 'ReportKpiChanged' -Default 'Changed'
        $lblAmbRows = Get-UiString 'ReportKpiAmbiguous' -Default 'Ambiguous'
        $lblAccRows = Get-UiString 'ReportKpiAccepted' -Default 'Accepted'

        [void]$html.AppendLine('    <div class="kpi-grid">')
        [void]$html.AppendLine("      <div class=""kpi-card""><div class=""num"">$($rows.Count)</div><div class=""lbl"">$lblAllRows</div></div>")
        [void]$html.AppendLine("      <div class=""kpi-card kpi-new""><div class=""num"">$kNew</div><div class=""lbl"">$lblNewRows</div></div>")
        [void]$html.AppendLine("      <div class=""kpi-card kpi-chg""><div class=""num"">$kChg</div><div class=""lbl"">$lblChgRows</div></div>")
        [void]$html.AppendLine("      <div class=""kpi-card kpi-amb""><div class=""num"">$kAmb</div><div class=""lbl"">$lblAmbRows</div></div>")
        [void]$html.AppendLine("      <div class=""kpi-card kpi-acc""><div class=""num"">$kAcc</div><div class=""lbl"">$lblAccRows</div></div>")
        [void]$html.AppendLine('    </div>')

        $secDetails = Get-UiString 'ReportSectionDetails' -Default 'Review Item Details'
        [void]$html.AppendLine("    <div class=""section-title"">$secDetails</div>")
        [void]$html.AppendLine('    <table>')
        [void]$html.AppendLine("      <thead><tr><th>#</th><th>$cStatus</th><th>$cDecision</th><th>$cId</th><th>$cBaseRow</th><th>$cSourceRow</th><th>$cChgSumm</th></tr></thead>")
        [void]$html.AppendLine('      <tbody>')

        $noticeNew  = Get-UiString 'ReportNewRecordNotice' -Default 'New record to append to base'
        $noticeNone = Get-UiString 'ReportNoChangesNotice' -Default 'No changes'

        foreach ($r in $rows) {
            $statusVal = $r.$cStatus
            $bClass = switch ($statusVal) {
                'New'       { 'b-new' }
                'Changed'   { 'b-chg' }
                'Ambiguous' { 'b-amb' }
                'Removed'   { 'b-rem' }
                default     { 'b-unch' }
            }
            $diffContent = if ($r.$cChgSumm) {
                [System.Net.WebUtility]::HtmlEncode($r.$cChgSumm)
            } elseif ($r.$cAmbReason) {
                "<span style='color:#F87171;'>$([System.Net.WebUtility]::HtmlEncode($r.$cAmbReason))</span>"
            } elseif ($statusVal -eq 'New') {
                "<span style='color:#34D399;'>$([System.Net.WebUtility]::HtmlEncode($noticeNew))</span>"
            } else {
                "<span style='color:#94A3B8;'>$([System.Net.WebUtility]::HtmlEncode($noticeNone))</span>"
            }

            [void]$html.AppendLine('        <tr>')
            [void]$html.AppendLine("          <td>$($r.$cIndex)</td>")
            [void]$html.AppendLine("          <td><span class=""badge $bClass"">$statusVal</span></td>")
            [void]$html.AppendLine("          <td>$($r.$cDecision)</td>")
            [void]$html.AppendLine("          <td><b>$([System.Net.WebUtility]::HtmlEncode($r.$cId))</b></td>")
            [void]$html.AppendLine("          <td>$($r.$cBaseRow)</td>")
            [void]$html.AppendLine("          <td>$($r.$cSourceRow)</td>")
            [void]$html.AppendLine("          <td>$diffContent</td>")
            [void]$html.AppendLine('        </tr>')
        }

        [void]$html.AppendLine('      </tbody>')
        [void]$html.AppendLine('    </table>')
        $reportFooter = Get-UiString 'ReportFooter' -Default 'Generated automatically by Excel Master Updater &bull; Zero External Dependencies'
        [void]$html.AppendLine("    <footer>$reportFooter</footer>")
        [void]$html.AppendLine('  </div>')
        [void]$html.AppendLine('</body>')
        [void]$html.AppendLine('</html>')

        [System.IO.File]::WriteAllText($DestinationPath, $html.ToString(), [System.Text.UTF8Encoding]::new($true))
    } elseif ($DestinationPath -match '\.xlsx$') {
        [FastExcelHelper]::ExportToExcel($DestinationPath, $rows, $true, $true, $true)
    } else {
        $csvLines = $rows | ConvertTo-Csv -NoTypeInformation -Delimiter ';'
        [System.IO.File]::WriteAllLines($DestinationPath, $csvLines, [System.Text.UTF8Encoding]::new($true))
    }

    return $DestinationPath
}

<#
.SYNOPSIS
    Executes automated headless comparison, reporting, and optional batch write-back (W3 & O2).

.DESCRIPTION
    Runs unattended synchronization between a master base file (.xlsx or .csv) and an incoming
    updates file (.xlsx or .csv). Automatically resolves mapping profile via header fingerprint,
    executes record comparison, classifies discrepancies (New, Changed, Ambiguous, Unchanged),
    applies the specified batch acceptance policy (-AutoAccept), performs pre-write backup,
    executes InPlace OpenXML cell patching or row appending, generates structured audit logs,
    and exports formatted discrepancy reports and machine-readable execution summary JSON.

.PARAMETER BaseFilePath
    Path to the master base Excel (.xlsx) or CSV (.csv) file to update. Mandatory.

.PARAMETER IncomingPath
    Path to the incoming change Excel (.xlsx) or CSV (.csv) file. Mandatory.

.PARAMETER BaseSheet
    Name of worksheet in base file. Defaults to first worksheet.

.PARAMETER IncomingSheet
    Name of worksheet in incoming file. Defaults to first worksheet.

.PARAMETER Config
    Optional hashtable containing configuration overrides.

.PARAMETER ConfigPath
    Optional path to custom config.json. Defaults to standard configuration path.

.PARAMETER AutoAccept
    Batch acceptance policy: 'None', 'AllNonAmbiguous', 'OnlyChanged', 'OnlyNew'.
    Defaults to 'None' (compare and report only; no write-back).

.PARAMETER ExportReportPath
    Optional file path to export review discrepancy report (.html, .xlsx, or .csv).

.PARAMETER SummaryJsonPath
    Optional file path to save execution summary JSON containing record counts and status metrics.
    Defaults to '<LogDirectory>\last_execution_summary.json' if not explicitly specified.

.PARAMETER Quiet
    Suppresses console host output.

.OUTPUTS
    [PSCustomObject] containing execution metrics:
    - Success (bool)
    - TotalIncoming (int)
    - NewRows (int)
    - ChangedRows (int)
    - AmbiguousRows (int)
    - UnchangedRows (int)
    - AcceptedRows (int)
    - UpdatedCells (int)
    - AddedRows (int)
    - BaseBackupPath (string)
    - ReportPath (string)
    - SummaryJsonPath (string)
    - BackupSizeWarning (string)
    - BaseSheetLog (bool)

.EXAMPLE
    $summary = Invoke-HeadlessMasterUpdater -BaseFilePath "Master.xlsx" -IncomingPath "Updates.csv" -AutoAccept AllNonAmbiguous -ExportReportPath "report.html" -SummaryJsonPath "summary.json"
#>
function Invoke-HeadlessMasterUpdater {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$BaseFilePath,

        [Parameter(Mandatory = $true)]
        [string]$IncomingPath,

        [string]$BaseSheet,
        [string]$IncomingSheet,
        [hashtable]$Config,
        [string]$ConfigPath,
        [ValidateSet('None', 'AllNonAmbiguous', 'OnlyChanged', 'OnlyNew')]
        [string]$AutoAccept = 'None',
        [string]$ExportReportPath,
        [string]$SummaryJsonPath,
        [switch]$Quiet
    )

    if (-not (Test-Path $BaseFilePath)) {
        throw "Base file not found: $BaseFilePath"
    }
    if (-not (Test-Path $IncomingPath)) {
        throw "Incoming file not found: $IncomingPath"
    }

    $baseFull = [System.IO.Path]::GetFullPath($BaseFilePath)
    $incFull  = [System.IO.Path]::GetFullPath($IncomingPath)
    if ($baseFull.Equals($incFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Base file and incoming file cannot be the same file: $BaseFilePath"
    }

    $baseExt = [System.IO.Path]::GetExtension($BaseFilePath).ToLowerInvariant()
    if ($baseExt -notin @('.xlsx', '.csv')) {
        throw "Unsupported base file format '$baseExt'. Only .xlsx and .csv files are supported."
    }
    $incExt = [System.IO.Path]::GetExtension($IncomingPath).ToLowerInvariant()
    if ($incExt -notin @('.xlsx', '.csv')) {
        throw "Unsupported incoming file format '$incExt'. Only .xlsx and .csv files are supported."
    }

    if ((Get-Item $BaseFilePath).Length -eq 0) {
        throw "Base file is empty (0 bytes): $BaseFilePath"
    }
    if ((Get-Item $IncomingPath).Length -eq 0) {
        throw "Incoming file is empty (0 bytes): $IncomingPath"
    }

    if ('ConsoleHelper' -as [type]) {
        try { [void][ConsoleHelper]::AttachConsole(-1) } catch { }
    }

    $cfg = if ($Config) { $Config } else {
        $effectiveCfgPath = if ([string]::IsNullOrWhiteSpace($ConfigPath)) { $script:DefaultConfigPath } else { $ConfigPath }
        Get-AppConfig -Path $effectiveCfgPath
    }

    if (-not $script:LanguagesCatalog -or $script:LanguagesCatalog.Count -eq 0) {
        if (Get-Command Import-LanguageCatalog -ErrorAction SilentlyContinue) {
            Import-LanguageCatalog -CatalogPath $cfg.LanguageCatalogPath
        }
    }

    # Sheets resolution
    if ([string]::IsNullOrEmpty($BaseSheet)) {
        $sheets = [FastExcelHelper]::GetSheetNames($BaseFilePath)
        $BaseSheet = if ($sheets -and $sheets.Count -gt 0) { $sheets[0] } else { 'Sheet1' }
    }
    if ([string]::IsNullOrEmpty($IncomingSheet)) {
        $sheets = [FastExcelHelper]::GetSheetNames($IncomingPath)
        $IncomingSheet = if ($sheets -and $sheets.Count -gt 0) { $sheets[0] } else { 'Sheet1' }
    }

    if (-not $Quiet) {
        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host " EXCEL MASTER UPDATER - HEADLESS BATCH PIPELINE" -ForegroundColor Cyan
        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host "Base File:     $BaseFilePath (Sheet: $BaseSheet)"
        Write-Host "Incoming File: $IncomingPath (Sheet: $IncomingSheet)"
        Write-Host "AutoAccept:    $AutoAccept"
    }

    $baseRaw = [FastExcelHelper]::ReadSheet($BaseFilePath, $BaseSheet)
    $baseHeaders = [FastExcelHelper]::GetHeaders($BaseFilePath, $BaseSheet)
    $incRaw = [FastExcelHelper]::ReadSheet($IncomingPath, $IncomingSheet)
    $incHeaders = [FastExcelHelper]::GetHeaders($IncomingPath, $IncomingSheet)
    # Stamp SourceFilePath so metadata tokens SourceFileFullPath / SourceFileName resolve correctly
    $incFullPath = [System.IO.Path]::GetFullPath($IncomingPath)
    foreach ($r in $incRaw) {
        if (-not $r.PSObject.Properties['SourceFilePath']) {
            $r | Add-Member -NotePropertyName 'SourceFilePath' -NotePropertyValue $incFullPath -Force
        } elseif ([string]::IsNullOrEmpty($r.SourceFilePath)) {
            $r.SourceFilePath = $incFullPath
        }
    }

    if ($incHeaders.Count -eq 0) {
        throw "Incoming file has no headers: $IncomingPath"
    }
    if ($baseHeaders.Count -eq 0) {
        throw "Base file has no headers: $BaseFilePath"
    }

    $fp = Compute-HeaderFingerprint -Headers $incHeaders -SheetName $IncomingSheet
    $fpFallback = Compute-HeaderFingerprint -Headers $incHeaders
    $profile = Get-MappingProfileByFingerprint -Fingerprint $fp -StorePath $cfg.ProfileStorePath
    if (-not $profile) {
        $profile = Get-MappingProfileByFingerprint -Fingerprint $fpFallback -StorePath $cfg.ProfileStorePath
    }

    $rules = [System.Collections.Generic.List[object]]::new()
    $joinBase = @()
    $joinInc  = @()

    if ($profile) {
        if (-not $Quiet) {
            Write-Host "Resolved mapping profile by fingerprint: $($profile.Name)" -ForegroundColor Green
        }
        $rules = $profile.MappingRules
        $joinBase = @($profile.JoinKeyBase)
        $joinInc  = @($profile.JoinKeyUpdate)
    } else {
        if (-not $Quiet) {
            Write-Host "No saved profile for fingerprint ($fp). Performing auto-mapping..." -ForegroundColor Yellow
        }
        foreach ($bCol in $baseHeaders) {
            $mInc = $incHeaders | Where-Object { $_.Trim().ToLowerInvariant() -eq $bCol.Trim().ToLowerInvariant() } | Select-Object -First 1
            if ($mInc) {
                $rules.Add([PSCustomObject]@{
                    BaseColumns   = @($bCol)
                    UpdateColumns = @($mInc)
                    MergeMode     = 'Exact'
                    BaseColumn    = $bCol
                    RuleType      = '1:1 Direct'
                    SourceColumns = @($mInc)
                    Delimiter     = ''
                    DefaultValue  = ''
                    CustomFormula = ''
                })
            }
        }
        $candidateKeys = @('ID', 'Pesel', 'Kod', 'Identyfikator', 'Key', 'Nr', 'Numer')
        foreach ($k in $candidateKeys) {
            $bk = $baseHeaders | Where-Object { $_.Trim().ToLowerInvariant() -eq $k.ToLowerInvariant() } | Select-Object -First 1
            $ik = $incHeaders  | Where-Object { $_.Trim().ToLowerInvariant() -eq $k.ToLowerInvariant() } | Select-Object -First 1
            if ($bk -and $ik) {
                $joinBase = @($bk)
                $joinInc  = @($ik)
                break
            }
        }
        if ($joinBase.Count -eq 0 -and $baseHeaders.Count -gt 0 -and $incHeaders.Count -gt 0) {
            $joinBase = @($baseHeaders[0])
            $joinInc  = @($incHeaders[0])
        }
    }

    $comp = @(Invoke-MasterCompare -BaseRows $baseRaw -IncomingRows $incRaw -MappingProfile $profile -MappingRules $rules -BaseJoinKey $joinBase -IncomingJoinKey $joinInc -BaseHeaders $baseHeaders -DetectRemoved ($cfg.DetectRemovedRows -eq $true) -MarkDeletedColumn $cfg.MarkDeletedColumn -MarkDeletedValue $cfg.MarkDeletedValue)

    $reviewItems = [System.Collections.Generic.List[object]]::new()
    $idx = 1
    foreach ($r in $comp) {
        $title = if ($r.IncomingRow -and $joinInc.Length -gt 0 -and $r.IncomingRow.Values -and $r.IncomingRow.Values.ContainsKey($joinInc[0])) {
            $r.IncomingRow.Values[$joinInc[0]].ToString()
        } elseif ($r.IncomingRow -and $joinInc.Length -gt 0 -and $r.IncomingRow.PSObject.Properties[$joinInc[0]]) {
            $r.IncomingRow.$($joinInc[0]).ToString()
        } elseif ($r.MatchedBaseRow -and $joinBase.Length -gt 0 -and $r.MatchedBaseRow.Values -and $r.MatchedBaseRow.Values.ContainsKey($joinBase[0])) {
            $r.MatchedBaseRow.Values[$joinBase[0]].ToString()
        } elseif ($r.MatchedBaseRow -and $joinBase.Length -gt 0 -and $r.MatchedBaseRow.PSObject.Properties[$joinBase[0]]) {
            $r.MatchedBaseRow.$($joinBase[0]).ToString()
        } else { "Row #$idx" }

        $dec = switch ($AutoAccept) {
            'AllNonAmbiguous' {
                if ($r.Status -in @('New', 'Changed', 'Removed')) { 'Accepted' } else { 'Skipped' }
            }
            'OnlyChanged' {
                if ($r.Status -eq 'Changed') { 'Accepted' } else { 'Skipped' }
            }
            'OnlyNew' {
                if ($r.Status -eq 'New') { 'Accepted' } else { 'Skipped' }
            }
            default { 'Pending' }
        }

        $reviewItems.Add([PSCustomObject]@{
            Record   = $r
            Index    = $idx
            IndexStr = "#$idx"
            Title    = $title
            Status   = $r.Status
            Decision = $dec
        })
        $idx++
    }

    $cNew  = @($comp | Where-Object { $_.Status -eq 'New' }).Count
    $cChg  = @($comp | Where-Object { $_.Status -eq 'Changed' }).Count
    $cUnc  = @($comp | Where-Object { $_.Status -eq 'Unchanged' }).Count
    $cAmb  = @($comp | Where-Object { $_.Status -eq 'Ambiguous' }).Count
    $cRem  = @($comp | Where-Object { $_.Status -eq 'Removed' }).Count
    $cAcc  = @($reviewItems | Where-Object { $_.Decision -eq 'Accepted' }).Count

    if (-not $Quiet) {
        Write-Host "Comparison Complete:" -ForegroundColor Cyan
        Write-Host "  Total: $($comp.Count) | New: $cNew | Changed: $cChg | Unchanged: $cUnc | Ambiguous: $cAmb | Removed: $cRem | Accepted: $cAcc"
    }

    $exportedReport = $null
    if ($ExportReportPath) {
        $exportedReport = Export-ReviewReport -ReviewItems $reviewItems -DestinationPath $ExportReportPath -BaseFilePath $BaseFilePath -IncomingPath $IncomingPath
        if (-not $Quiet) {
            Write-Host "Exported diff report to: $exportedReport" -ForegroundColor Green
        }
    }

    $wbRes = $null
    $logRes = $null

    if ($cAcc -gt 0) {
        $accRecords = [System.Collections.Generic.List[object]]::new()
        foreach ($it in $reviewItems) {
            if ($it.Decision -eq 'Accepted') {
                $accRecords.Add($it.Record)
            }
        }
        try {
            $wbRes = Invoke-MasterWriteBack -BaseFilePath $BaseFilePath -BaseSheet $BaseSheet -ReviewItems $accRecords -AppConfig $cfg -BaseHeaders $baseHeaders
            $logRes = Write-ImportLog -LogDirectory $cfg.LogDirectory -BatchId $wbRes.BatchId -ReviewItems $reviewItems -WriteBackResult $wbRes -BaseFilePath $BaseFilePath -RedactNames $cfg.RedactNamesInLog -LogChangesToBaseSheet ($cfg.LogChangesToBaseSheet -eq $true) -BaseSheetLogName ($(if ($cfg.BaseSheetLogName) { $cfg.BaseSheetLogName } else { 'ImportLog' }))
        } catch {
            $errMsg = $_.Exception.Message
            if (-not $Quiet) {
                Write-Host "  ERROR during write-back: $errMsg" -ForegroundColor Red
            }
            try {
                $masterJsonlPath = Join-Path $cfg.LogDirectory "masterlog.jsonl"
                $failEvt = [ordered]@{
                    ts         = (Get-Date).ToUniversalTime().ToString('o')
                    action     = 'WriteBackFailed'
                    baseFile   = $BaseFilePath
                    incoming   = $IncomingPath
                    error      = $errMsg
                    user       = $env:USERNAME
                }
                [System.IO.File]::AppendAllText($masterJsonlPath, (($failEvt | ConvertTo-Json -Compress) + "`r`n"), [System.Text.UTF8Encoding]::new($true))
            } catch { }
            throw
        }

        if (-not $Quiet) {
            Write-Host "Write-Back Complete:" -ForegroundColor Green
            Write-Host "  Updated Cells: $($wbRes.UpdatedCells) | Added Rows: $($wbRes.AddedRows)"
            Write-Host "  Backup: $($wbRes.BackupPath)"
            Write-Host "  Log:    $($logRes.TxtPath)"
            if ($logRes.BaseSheetLog) {
                Write-Host "  Sheet Log: $($logRes.BaseSheetLog)"
            }
            if ($wbRes.BackupSizeWarning) {
                Write-Host "  [BACKUP WARNING] $($wbRes.BackupSizeWarning)" -ForegroundColor Yellow
            }
        }
    } else {
        if (-not $Quiet) {
            Write-Host "No records accepted for write-back. Base file untouched." -ForegroundColor Yellow
        }
    }

    $result = [PSCustomObject]@{
        Success           = $true
        TotalIncoming     = $comp.Count
        NewRows           = $cNew
        ChangedRows       = $cChg
        UnchangedRows     = $cUnc
        AmbiguousRows     = $cAmb
        RemovedRows       = $cRem
        AcceptedRows      = $cAcc
        UpdatedCells      = if ($wbRes) { $wbRes.UpdatedCells } else { 0 }
        AddedRows         = if ($wbRes) { $wbRes.AddedRows } else { 0 }
        BackupPath        = if ($wbRes) { $wbRes.BackupPath } else { $null }
        LogJsonlPath      = if ($logRes) { $logRes.JsonlPath } else { $null }
        LogTxtPath        = if ($logRes) { $logRes.TxtPath } else { $null }
        ReportPath        = $exportedReport
        BackupSizeWarning = if ($wbRes) { $wbRes.BackupSizeWarning } else { $null }
        BaseSheetLog      = if ($logRes) { $logRes.BaseSheetLog } else { $null }
    }

    $targetSummaryPath = if ($SummaryJsonPath) {
        $SummaryJsonPath
    } elseif ($cfg.LogDirectory -and (Test-Path $cfg.LogDirectory)) {
        Join-Path $cfg.LogDirectory "last_execution_summary.json"
    } else {
        $null
    }
    if ($targetSummaryPath) {
        try {
            $summaryJson = $result | ConvertTo-Json -Depth 5
            [System.IO.File]::WriteAllText($targetSummaryPath, $summaryJson, [System.Text.UTF8Encoding]::new($true))
            $result | Add-Member -NotePropertyName 'SummaryJsonPath' -NotePropertyValue $targetSummaryPath -Force
        } catch { }
    }

    $global:LASTEXITCODE = 0
    return $result
}

# ==============================================================================
# Region 9: WPF UI & Localization Catalog (PL / EN / DE)
# ==============================================================================
$script:LanguagesCatalog = [ordered]@{}

<#
.SYNOPSIS
    Loads external language catalog with built-in fallback.

.DESCRIPTION
    Reads language.json containing trilingual strings (PL, EN, DE). Falls back to embedded
    dictionaries if the file cannot be accessed.

.PARAMETER CatalogPath
    Optional path to language.json.

.OUTPUTS
    System.Void.
#>
function Import-LanguageCatalog {
    param([string]$CatalogPath)
    $script:LanguagesCatalog = [ordered]@{}

    # 1. Try reading external language.json
    $foundPath = $null
    $candidates = @(
        $CatalogPath,
        (Join-Path $PSScriptRoot 'language.json'),
        (Join-Path (Split-Path -Parent $PSScriptRoot) 'language.json'),
        (Join-Path (Get-Location).Path 'language.json')
    )
    foreach ($cand in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace($cand) -and (Test-Path -LiteralPath $cand)) {
            $foundPath = $cand
            break
        }
    }

    if ($foundPath) {
        try {
            $raw = [System.IO.File]::ReadAllText($foundPath, [System.Text.Encoding]::UTF8)
            $parsed = $raw | ConvertFrom-Json
            foreach ($prop in $parsed.Languages.PSObject.Properties) {
                $code = $prop.Name.ToLower()
                $strMap = @{}
                foreach ($sProp in $prop.Value.Strings.PSObject.Properties) {
                    $strMap[$sProp.Name] = [string]$sProp.Value
                }
                $script:LanguagesCatalog[$code] = [PSCustomObject]@{
                    Code        = $code
                    DisplayName = [string]$prop.Value.DisplayName
                    Strings     = $strMap
                }
            }
        } catch {
            Write-Warning "Could not parse language catalog $foundPath : $_"
        }
    }

    # 2. Ensure default languages exist
    if (-not $script:LanguagesCatalog.Contains('pl')) {
        $script:LanguagesCatalog['pl'] = [PSCustomObject]@{
            Code        = 'pl'
            DisplayName = 'Polski'
            Strings     = @{
                AppTitle            = 'Excel Master Updater'
                BadgeProfileAuto    = 'Profil: Auto'
                BadgeProfileManual  = 'Profil: Ręczny'
                BadgeProfileNone    = 'Profil: Brak'
                LblBasePath         = 'Plik bazy:'
                LblIncomingPath     = 'Plik zmian:'
                LblSheet            = 'Arkusz:'
                BtnBrowseBase       = 'Wybierz bazę...'
                BtnBrowseIncoming   = 'Wybierz zmiany...'
                BtnThemeDark        = '🌙 Ciemny'
                BtnThemeLight       = '☀ Jasny'
                BtnSettings         = '⚙ Ustawienia'
                LblLanguage         = 'Język:'
                TabMapping          = '1. Mapowanie kolumn'
                BtnAutoMap          = '⚡ Automatyczne mapowanie'
                BtnAddRule          = '+ Dodaj regułę'
                BtnRemoveRule       = 'Usuń regułę'
                BtnSaveProfile      = 'Zapisz profil'
                ColBase             = 'Kolumny w bazie'
                ColIncoming         = 'Kolumny w pliku zmian'
                ColMergeMode        = 'Tryb łączenia'
                ColSeparator        = 'Separator'
                LblJoinKeyTitle     = 'Klucz złączenia (Join Key):'
                LblJoinBase         = 'Kolumny klucza bazy:'
                LblJoinIncoming     = 'Kolumny klucza zmian:'
                BtnRunCompare       = 'Rozpocznij porównanie ▶'
                TabReview           = '2. Przegląd i zatwierdzanie'
                SearchPlaceholder   = 'Szukaj wierszy...'
                FilterAll           = 'Wszystkie'
                FilterNew           = 'Nowe'
                FilterChanged       = 'Zmienione'
                FilterAmbiguous     = 'Niejednoznaczne'
                FilterAccepted      = 'Zaakceptowane'
                FilterSkipped       = 'Pominięte'
                FilterRejected      = 'Odrzucone'
                CountersFormat      = 'Wiersze: {0} | Nowe: {1} | Zmienione: {2} | Niejednoznaczne: {3} | Pominięte: {4} | Zaakceptowane: {5} | Odrzucone: {6}'
                SelectRowHeader     = 'Wybierz wiersz z listy po lewej stronie'
                BtnAccept           = '✔ Akceptuj (A)'
                BtnReject           = '✖ Odrzuć (R)'
                BtnSkip             = '⏭ Pomiń (S)'
                BtnEdit             = '✏ Zmień wartość... (E)'
                BtnTreatAsNew       = '➕ Traktuj jako nowy wiersz'
                BtnApplyAccepted    = 'Zastosuj zaakceptowane zmiany do bazy'
                StatusReady         = 'Gotowy do pracy.'
                StatusComparing     = 'Porównywanie rekordów...'
                StatusCompareDone   = 'Porównanie zakończone: {0} wierszy do przeglądu.'
                StatusWriting       = 'Zapisywanie zmian do bazy...'
                StatusWriteSuccess  = 'Pomyślnie zaktualizowano bazę! Kopia zapasowa utworzona.'
                StatusWriteError    = 'Błąd podczas zapisu bazy: {0}'
                StatusNoAccepted    = 'Brak zaakceptowanych zmian do zapisania.'
                ConfirmApplyTitle   = 'Potwierdzenie zapisu'
                ConfirmApplyMsg     = "Czy zastosować {0} zaakceptowanych zmian do pliku bazy?\nKopia zapasowa zostanie utworzona automatycznie."
                ProfileSavedTitle   = 'Zapisano profil'
                ProfileSavedMsg     = "Profil '{0}' został pomyślnie zapisany."
                ProfileNamePrompt   = 'Podaj nazwę profilu:'
                ErrorTitle          = 'Błąd'
                WarningTitle        = 'Ostrzeżenie'
                InfoTitle           = 'Informacja'
                ColField            = 'Pole'
                ColCurrentBase      = 'Bieżąca wartość w bazie'
                ColIncomingValue    = 'Nowa wartość ze zmian'
                ColApply            = 'Zmień'
                StatusBadgeNew      = 'NOWY'
                StatusBadgeChanged  = 'ZMIENIONY'
                StatusBadgeAmbiguous= 'NIEJEDNOZNACZNY'
                StatusBadgeAccepted = 'ZAAKCEPTOWANY'
                StatusBadgeRejected = 'ODRZUCONY'
                StatusBadgeSkipped  = 'POMINIĘTY'
                RowNumberFormat     = 'Wiersz {0}'
                AmbiguousHeader     = 'Dopasowanie niejednoznaczne: znaleziono wielu kandydatów w bazie'
                AmbiguousSelectPrompt = 'Wybierz właściwy rekord z bazy lub potraktuj jako nowy wiersz:'
                EditPromptTitle     = 'Edycja wartości'
                EditPromptMsg       = "Podaj nową wartość dla '{0}':"
                SettingsTitle       = 'Ustawienia'
                SettingsBackupCount = 'Liczba kopii zapasowych:'
                SettingsWriteMode   = 'Tryb zapisu:'
                SettingsMaskPii     = 'Maskuj dane osobowe (PII) w logach'
                SettingsBtnSave     = 'Zapisz ustawienia'
                SettingsBtnCancel   = 'Anuluj'
                StatusBadgeRemoved  = 'USUNIĘTY'
                SettingsDetectRemoved = 'Oznacz brakujące rekordy jako usunięte'
                FilterRemoved       = 'Usunięte'
                BtnBack             = '◀ Wstecz (B)'
                BtnRestoreBackup    = '↺ Przywróć kopię'
                RestoreConfirmTitle = 'Przywracanie bazy z kopii'
                RestoreConfirmMsg   = "Czy na pewno chcesz przywrócić plik bazy z najnowszej kopii zapasowej:
{0} ?"
                RestoreNoBackups    = 'Nie znaleziono kopii zapasowych dla tego pliku bazy.'
                RestoreSuccess      = 'Pomyślnie przywrócono bazę z kopii zapasowej!'
                BtnAcceptAll        = '✔✔ Akceptuj wszystkie'
                BtnRejectAll        = '✖✖ Odrzuć wszystkie'
                BtnExportReport     = '📊 Eksportuj raport...'
                ExportSuccess       = "Pomyślnie wyeksportowano raport do:
{0}"
                AcceptAllDone       = 'Zaakceptowano {0} rekordów.'
                RejectAllDone       = 'Odrzucono {0} rekordów.'
            }
        }
    }
}

Import-LanguageCatalog

<#
.SYNOPSIS
    Retrieves localized UI text for a key.

.DESCRIPTION
    Looks up translation in active language catalog, falling back to Polish or the key name itself.

.PARAMETER Key
    Catalog string key identifier.

.OUTPUTS
    System.String. Localized string.
#>
function Get-UiString {
    param([string]$Key, [string]$Default = '')
    $lang = if ($script:CurrentLanguage) { $script:CurrentLanguage.ToLower() } else { 'pl' }
    if ($script:LanguagesCatalog -and $script:LanguagesCatalog.Contains($lang)) {
        $dict = $script:LanguagesCatalog[$lang].Strings
        if ($dict -and $dict.ContainsKey($Key)) {
            return $dict[$Key]
        }
    }
    # Fallback to English
    if ($script:LanguagesCatalog -and $script:LanguagesCatalog.Contains('en')) {
        $dict = $script:LanguagesCatalog['en'].Strings
        if ($dict -and $dict.ContainsKey($Key)) {
            return $dict[$Key]
        }
    }
    if (-not [string]::IsNullOrEmpty($Default)) { return $Default }
    return $Key
}

<#
.SYNOPSIS
    Returns color palette brush definitions for UI theme.

.DESCRIPTION
    Provides curated hex color palettes for Dark and Light modes.

.PARAMETER ThemeName
    Theme name ('Dark' or 'Light').

.OUTPUTS
    Hashtable of hex color strings.
#>
function Get-UpdaterThemePalette([string]$ThemeName) {
    if ($ThemeName -eq 'Dark') {
        return [ordered]@{
            IsDark             = $true
            BgApp              = '#0F172A'   # Slate 900
            BgWindow           = '#0F172A'   # Legacy alias
            BgHeader           = '#1E293B'   # Slate 800
            BgCard             = '#1E293B'   # Slate 800
            BgCardHover        = '#293548'   # Slate 700
            BgCardAlt          = '#162032'   # Slate 850
            BorderCard         = '#334155'   # Slate 700
            BorderSubtle       = '#1E293B'   # Slate 800
            BorderInput        = '#374151'   # Gray 700
            TextPrimary        = '#F8FAFC'   # Slate 50
            TextSecondary      = '#94A3B8'   # Slate 400
            TextMuted          = '#64748B'   # Slate 500
            BgInput            = '#111827'   # Gray 900
            BtnSecondaryBg     = '#334155'   # Slate 700
            BgButtonDefault    = '#334155'   # Legacy alias
            BtnSecondaryFg     = '#F8FAFC'   # Slate 50
            BtnSecondaryHover  = '#475569'   # Slate 600
            BgButtonHover      = '#475569'   # Legacy alias
            AccentBlue         = '#2563EB'   # Royal Blue
            BgAccent           = '#2563EB'   # Legacy alias
            AccentBlueHover    = '#1D4ED8'
            AccentGreen        = '#059669'   # Emerald
            AccentAmber        = '#D97706'   # Amber
            AccentRed          = '#DC2626'   # Red
            GridLines          = '#2D3748'
            DgGridBrush        = '#2D3748'   # Alias for GridLines
            DataGridHeaderBg   = '#111827'
            DataGridHeaderFg   = '#94A3B8'
            DataGridRowBg      = '#1E293B'
            DgRowBg            = '#1E293B'   # Alias for DataGridRowBg
            DataGridAltRowBg   = '#162032'
            DgAltRowBg         = '#162032'   # Alias for DataGridAltRowBg
            StatusBarBg        = '#0A0F1D'
            StatusBarFg        = '#94A3B8'
            # Badges
            RowNewBg           = '#064E3B'   # Emerald 900
            RowNewFg           = '#A7F3D0'   # Emerald 200
            RowNewBorder       = '#059669'
            RowChangedBg       = '#451A03'   # Amber 950
            RowChangedFg       = '#FDE68A'   # Amber 200
            RowChangedBorder   = '#D97706'
            RowAmbiguousBg     = '#172554'   # Blue 950
            RowAmbiguousFg     = '#BFDBFE'   # Blue 200
            RowAmbiguousBorder = '#3B82F6'
            RowRemovedBg       = '#4C0519'   # Rose 950
            RowRemovedFg       = '#FECDD3'   # Rose 200
            RowRemovedBorder   = '#E11D48'
            RowAcceptedBg      = '#1E293B'   # Slate 800
            RowAcceptedFg      = '#94A3B8'   # Slate 400
            RowAcceptedBorder  = '#334155'
            RowRejectedBg      = '#450A0A'   # Red 950
            RowRejectedFg      = '#FECACA'   # Red 200
            RowRejectedBorder  = '#DC2626'
            RowSkippedBg       = '#1E293B'   # Slate 800
            RowSkippedFg       = '#64748B'   # Slate 500
            RowSkippedBorder   = '#334155'
            # Diff cells
            DiffCellNewBg      = '#3B1E08'
            DiffCellNewFg      = '#FDE68A'
            DiffCellNewBorder  = '#F59E0B'
            CellChgBg          = '#3B1E08'   # Legacy alias
            CellChgBorder      = '#F59E0B'   # Legacy alias
            DiffCellOldFg      = '#94A3B8'
            AlertAmbBg         = '#172554'
            AlertAmbBorder     = '#3B82F6'
            AlertAmbHdrFg      = '#BFDBFE'
            AlertAmbSubFg      = '#93C5FD'
        }
    } else {
        return [ordered]@{
            IsDark             = $false
            BgApp              = '#F1F5F9'   # Slate 100
            BgWindow           = '#F1F5F9'   # Legacy alias
            BgHeader           = '#FFFFFF'   # White
            BgCard             = '#FFFFFF'   # White
            BgCardHover        = '#F8FAFC'   # Slate 50
            BgCardAlt          = '#F8FAFC'   # Slate 50
            BorderCard         = '#CBD5E1'   # Slate 300
            BorderSubtle       = '#E2E8F0'   # Slate 200
            BorderInput        = '#CBD5E1'   # Slate 300
            TextPrimary        = '#0F172A'   # Slate 900
            TextSecondary      = '#475569'   # Slate 600
            TextMuted          = '#64748B'   # Slate 500
            BgInput            = '#FFFFFF'   # White
            BtnSecondaryBg     = '#E2E8F0'   # Slate 200
            BgButtonDefault    = '#E2E8F0'   # Legacy alias
            BtnSecondaryFg     = '#0F172A'   # Slate 900
            BtnSecondaryHover  = '#CBD5E1'   # Slate 300
            BgButtonHover      = '#CBD5E1'   # Legacy alias
            AccentBlue         = '#2563EB'   # Royal Blue
            BgAccent           = '#2563EB'   # Legacy alias
            AccentBlueHover    = '#1D4ED8'
            AccentGreen        = '#059669'   # Emerald
            AccentAmber        = '#D97706'   # Amber
            AccentRed          = '#DC2626'   # Red
            GridLines          = '#E2E8F0'
            DgGridBrush        = '#E2E8F0'   # Alias for GridLines
            DataGridHeaderBg   = '#E2E8F0'
            DataGridHeaderFg   = '#334155'
            DataGridRowBg      = '#FFFFFF'
            DgRowBg            = '#FFFFFF'   # Alias for DataGridRowBg
            DataGridAltRowBg   = '#F8FAFC'
            DgAltRowBg         = '#F8FAFC'   # Alias for DataGridAltRowBg
            StatusBarBg        = '#E2E8F0'
            StatusBarFg        = '#475569'
            # Badges
            RowNewBg           = '#DCFCE7'   # Emerald 100
            RowNewFg           = '#166534'   # Emerald 800
            RowNewBorder       = '#86EFAC'
            RowChangedBg       = '#FEF3C7'   # Amber 100
            RowChangedFg       = '#92400E'   # Amber 800
            RowChangedBorder   = '#FCD34D'
            RowAmbiguousBg     = '#DBEAFE'   # Blue 100
            RowAmbiguousFg     = '#1E40AF'   # Blue 800
            RowAmbiguousBorder = '#93C5FD'
            RowRemovedBg       = '#FFE4E6'   # Rose 100
            RowRemovedFg       = '#9F1239'   # Rose 800
            RowRemovedBorder   = '#FDA4AF'
            RowAcceptedBg      = '#F1F5F9'   # Slate 100
            RowAcceptedFg      = '#475569'   # Slate 600
            RowAcceptedBorder  = '#CBD5E1'
            RowRejectedBg      = '#FEE2E2'   # Red 100
            RowRejectedFg      = '#991B1B'   # Red 800
            RowRejectedBorder  = '#FCA5A5'
            RowSkippedBg       = '#E2E8F0'   # Slate 200
            RowSkippedFg       = '#64748B'   # Slate 500
            RowSkippedBorder   = '#CBD5E1'
            # Diff cells
            DiffCellNewBg      = '#FEF3C7'   # Amber 100
            DiffCellNewFg      = '#92400E'   # Amber 800
            DiffCellNewBorder  = '#F59E0B'
            CellChgBg          = '#FEF3C7'   # Legacy alias
            CellChgBorder      = '#F59E0B'   # Legacy alias
            DiffCellOldFg      = '#475569'
            AlertAmbBg         = '#EFF6FF'   # Blue 50
            AlertAmbBorder     = '#3B82F6'
            AlertAmbHdrFg      = '#1E40AF'
            AlertAmbSubFg      = '#2563EB'
        }
    }
}

function Get-StatusBadgeColors {
    param(
        [string]$Status,
        [string]$ThemeName = 'Dark'
    )
    $p = Get-UpdaterThemePalette $ThemeName
    switch ($Status) {
        'New'       { return @{ Bg = $p.RowNewBg;       Fg = $p.RowNewFg;       Border = $p.RowNewBorder } }
        'Changed'   { return @{ Bg = $p.RowChangedBg;   Fg = $p.RowChangedFg;   Border = $p.RowChangedBorder } }
        'Removed'   { return @{ Bg = $p.RowRemovedBg;   Fg = $p.RowRemovedFg;   Border = $p.RowRemovedBorder } }
        'Ambiguous' { return @{ Bg = $p.RowAmbiguousBg; Fg = $p.RowAmbiguousFg; Border = $p.RowAmbiguousBorder } }
        'Accepted'  { return @{ Bg = $p.RowAcceptedBg;  Fg = $p.RowAcceptedFg;  Border = $p.RowAcceptedBorder } }
        'Rejected'  { return @{ Bg = $p.RowRejectedBg;  Fg = $p.RowRejectedFg;  Border = $p.RowRejectedBorder } }
        'Skipped'   { return @{ Bg = $p.RowSkippedBg;   Fg = $p.RowSkippedFg;   Border = $p.RowSkippedBorder } }
        'Unchanged' { return @{ Bg = $p.RowSkippedBg;   Fg = $p.RowSkippedFg;   Border = $p.RowSkippedBorder } }
        default     { return @{ Bg = $p.RowSkippedBg;   Fg = $p.RowSkippedFg;   Border = $p.RowSkippedBorder } }
    }
}

# ==============================================================================
# Region 10: Main Entry Point (Show-MasterUpdater)
# ==============================================================================
<#
.SYNOPSIS
    Launches the Excel Master Updater graphical user interface.

.DESCRIPTION
    Builds and renders the full WPF desktop window, wires up events, theming, localization,
    and review list controllers.

.PARAMETER InitBaseFilePath
    Optional initial base file path to load on startup.

.PARAMETER InitIncomingPath
    Optional initial incoming file path to load on startup.

.PARAMETER InitBaseSheet
    Optional initial base sheet name.

.PARAMETER InitIncomingSheet
    Optional initial incoming sheet name.

.PARAMETER CustomConfigPath
    Optional custom config.json path.

.PARAMETER NonInteractive
    Runs window initialization without displaying modal dialog (headless verification).

.OUTPUTS
    System.Void.
#>
function Show-MasterUpdater {
    param(
        [string]$InitBaseFilePath,
        [string]$InitIncomingPath,
        [string]$InitBaseSheet,
        [string]$InitIncomingSheet,
        [string]$CustomConfigPath,
        [switch]$NonInteractive
    )

    $cfgPath = if ([string]::IsNullOrWhiteSpace($CustomConfigPath)) { $script:DefaultConfigPath } else { $CustomConfigPath }
    $cfg = Get-AppConfig -Path $cfgPath

    if (-not [string]::IsNullOrEmpty($InitBaseFilePath)) { $cfg.BaseFilePath = $InitBaseFilePath }
    if (-not [string]::IsNullOrEmpty($InitBaseSheet)) { $cfg.BaseSheet = $InitBaseSheet }

    # Load Language Catalog before user-facing prompts
    Import-LanguageCatalog
    $script:CurrentLanguage = if ($cfg.UiLanguage) { $cfg.UiLanguage.ToLower() } else { 'pl' }

    # Detect crash mid-write (.tmp.xlsx recovery)
    if (-not [string]::IsNullOrEmpty($cfg.BaseFilePath)) {
        $tmpFile = "$($cfg.BaseFilePath).tmp.xlsx"
        if (Test-Path $tmpFile) {
            if (-not $NonInteractive) {
                $msg = (Get-UiString 'CrashRecoveryMsg') -f $tmpFile
                $res = [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'CrashRecoveryTitle'), [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                if ($res -eq [System.Windows.Forms.DialogResult]::Yes) {
                    try { Remove-Item -Force $tmpFile } catch { }
                }
            } else {
                try { Remove-Item -Force $tmpFile } catch { }
            }
        }
    }

    # Build XAML UI Window
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        x:Name="winMasterUpdater"
        Title="Excel Master Updater"
        MinWidth="1100" MinHeight="720"
        Width="1280" Height="860"
        WindowStartupLocation="CenterScreen"
        FontFamily="Segoe UI" FontSize="13"
        TextOptions.TextFormattingMode="Display"
        TextOptions.TextRenderingMode="ClearType"
        SnapsToDevicePixels="True"
        UseLayoutRounding="True">
    <Window.Resources>
        <!-- Dynamic Palette Brushes -->
        <SolidColorBrush x:Key="BgApp" Color="#0F172A"/>
        <SolidColorBrush x:Key="BgHeader" Color="#1E293B"/>
        <SolidColorBrush x:Key="BgCard" Color="#1E293B"/>
        <SolidColorBrush x:Key="BgCardHover" Color="#293548"/>
        <SolidColorBrush x:Key="BorderCard" Color="#334155"/>
        <SolidColorBrush x:Key="BorderSubtle" Color="#1E293B"/>
        <SolidColorBrush x:Key="BorderInput" Color="#374151"/>
        <SolidColorBrush x:Key="TextPrimary" Color="#F8FAFC"/>
        <SolidColorBrush x:Key="TextSecondary" Color="#94A3B8"/>
        <SolidColorBrush x:Key="TextMuted" Color="#64748B"/>
        <SolidColorBrush x:Key="BgInput" Color="#111827"/>
        <SolidColorBrush x:Key="BtnSecondaryBg" Color="#334155"/>
        <SolidColorBrush x:Key="BtnSecondaryFg" Color="#F8FAFC"/>
        <SolidColorBrush x:Key="AccentBlue" Color="#2563EB"/>
        <SolidColorBrush x:Key="AccentGreen" Color="#059669"/>
        <SolidColorBrush x:Key="AccentAmber" Color="#D97706"/>
        <SolidColorBrush x:Key="AccentRed" Color="#DC2626"/>
        <SolidColorBrush x:Key="DataGridRowBg" Color="#1E293B"/>
        <SolidColorBrush x:Key="DataGridAltRowBg" Color="#162032"/>
        <SolidColorBrush x:Key="GridLines" Color="#2D3748"/>
        <SolidColorBrush x:Key="DataGridHeaderBg" Color="#111827"/>
        <SolidColorBrush x:Key="DataGridHeaderFg" Color="#94A3B8"/>
        <SolidColorBrush x:Key="StatusBarBg" Color="#0A0F1D"/>
        <SolidColorBrush x:Key="StatusBarFg" Color="#94A3B8"/>

        <!-- Modern Flat Button Style & ControlTemplate -->
        <Style TargetType="Button">
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Padding" Value="14,6"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderCard}"/>
            <Setter Property="Background" Value="{DynamicResource BtnSecondaryBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource BtnSecondaryFg}"/>
            <Setter Property="SnapsToDevicePixels" Value="True"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="btnBdr" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6" Padding="{TemplateBinding Padding}" SnapsToDevicePixels="True">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" RecognizesAccessKey="True"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="btnBdr" Property="Opacity" Value="0.88"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="btnBdr" Property="Opacity" Value="0.75"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="btnBdr" Property="Opacity" Value="0.4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Modern Flat TextBox Style & ControlTemplate -->
        <Style TargetType="TextBox">
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Background" Value="{DynamicResource BgInput}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderInput}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="CaretBrush" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="SnapsToDevicePixels" Value="True"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border x:Name="tbBdr" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6" Padding="{TemplateBinding Padding}" SnapsToDevicePixels="True">
                            <ScrollViewer x:Name="PART_ContentHost" Focusable="False" HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Hidden"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="tbBdr" Property="BorderBrush" Value="{DynamicResource AccentBlue}"/>
                            </Trigger>
                            <Trigger Property="IsKeyboardFocused" Value="True">
                                <Setter TargetName="tbBdr" Property="BorderBrush" Value="{DynamicResource AccentBlue}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Modern ComboBoxItem Style -->
        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="{DynamicResource BgCard}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border x:Name="cbiBorder" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}" SnapsToDevicePixels="True">
                            <ContentPresenter/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="cbiBorder" Property="Background" Value="{DynamicResource BgCardHover}"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="cbiBorder" Property="Background" Value="{DynamicResource AccentBlue}"/>
                                <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Modern Flat ComboBox Style & ControlTemplate -->
        <ControlTemplate x:Key="ComboBoxToggleButton" TargetType="ToggleButton">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition />
                    <ColumnDefinition Width="28" />
                </Grid.ColumnDefinitions>
                <Border x:Name="Border" Grid.ColumnSpan="2" CornerRadius="6"
                        Background="{TemplateBinding Background}"
                        BorderBrush="{TemplateBinding BorderBrush}"
                        BorderThickness="{TemplateBinding BorderThickness}" />
                <Path x:Name="Arrow" Grid.Column="1" HorizontalAlignment="Center" VerticalAlignment="Center"
                      Data="M 0 0 L 4 4 L 8 0 Z" Fill="{DynamicResource TextSecondary}" />
            </Grid>
            <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="true">
                    <Setter TargetName="Border" Property="BorderBrush" Value="{DynamicResource AccentBlue}" />
                </Trigger>
                <Trigger Property="IsEnabled" Value="false">
                    <Setter TargetName="Border" Property="Opacity" Value="0.5" />
                    <Setter TargetName="Arrow" Property="Opacity" Value="0.5" />
                </Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate>

        <Style TargetType="ComboBox">
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Background" Value="{DynamicResource BgInput}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderInput}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.VerticalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.CanContentScroll" Value="True"/>
            <Setter Property="SnapsToDevicePixels" Value="True"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton Name="ToggleButton"
                                          Template="{StaticResource ComboBoxToggleButton}"
                                          Focusable="false"
                                          IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                          ClickMode="Press"
                                          Background="{TemplateBinding Background}"
                                          BorderBrush="{TemplateBinding BorderBrush}"
                                          BorderThickness="{TemplateBinding BorderThickness}"/>
                            <ContentPresenter Name="ContentSite"
                                              IsHitTestVisible="False"
                                              Content="{TemplateBinding SelectionBoxItem}"
                                              ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                              ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"
                                              Margin="10,6,28,6"
                                              VerticalAlignment="Center"
                                              HorizontalAlignment="Left"/>
                            <Popup Name="Popup"
                                   Placement="Bottom"
                                   IsOpen="{TemplateBinding IsDropDownOpen}"
                                   AllowsTransparency="True"
                                   Focusable="False"
                                   PopupAnimation="Slide">
                                <Grid Name="DropDown"
                                      SnapsToDevicePixels="True"
                                      MinWidth="{TemplateBinding ActualWidth}"
                                      MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                    <Border x:Name="DropDownBorder"
                                            Background="{DynamicResource BgCard}"
                                            BorderThickness="1"
                                            BorderBrush="{DynamicResource BorderCard}"
                                            CornerRadius="6"
                                            Margin="0,2,0,4">
                                        <ScrollViewer Margin="2,4" SnapsToDevicePixels="True">
                                            <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained" />
                                        </ScrollViewer>
                                    </Border>
                                </Grid>
                            </Popup>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="HasItems" Value="false">
                                <Setter TargetName="DropDownBorder" Property="MinHeight" Value="30"/>
                            </Trigger>
                            <Trigger Property="IsGrouping" Value="true">
                                <Setter Property="ScrollViewer.CanContentScroll" Value="false"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Modern Flat TabControl & TabItem Style -->
        <Style TargetType="TabControl">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabControl">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="0,0,0,1" Padding="16,0">
                                <TabPanel IsItemsHost="True" VerticalAlignment="Center"/>
                            </Border>
                            <ContentPresenter Grid.Row="1" ContentSource="SelectedContent" Margin="{TemplateBinding Padding}"/>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="TabItem">
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="{DynamicResource TextSecondary}"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Padding" Value="18,12"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Grid x:Name="root">
                            <Border x:Name="tabBorder" Background="Transparent" Padding="{TemplateBinding Padding}">
                                <ContentPresenter ContentSource="Header" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                            <Border x:Name="activeIndicator" Height="3" Background="Transparent" VerticalAlignment="Bottom" Margin="4,0,4,0"/>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter Property="Foreground" Value="{DynamicResource AccentBlue}"/>
                                <Setter TargetName="activeIndicator" Property="Background" Value="{DynamicResource AccentBlue}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Modern ListBox & ListBoxItem Style -->
        <Style TargetType="ListBox">
            <Setter Property="Background" Value="{DynamicResource BgCard}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderCard}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Disabled"/>
            <Setter Property="ScrollViewer.CanContentScroll" Value="True"/>
            <Setter Property="VirtualizingStackPanel.IsVirtualizing" Value="True"/>
            <Setter Property="VirtualizingStackPanel.VirtualizationMode" Value="Recycling"/>
        </Style>

        <Style TargetType="ListBoxItem">
            <Setter Property="Padding" Value="0"/>
            <Setter Property="Margin" Value="0,0,0,4"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBoxItem">
                        <Border x:Name="itemBorder" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Padding="10,8" SnapsToDevicePixels="True">
                            <ContentPresenter/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="itemBorder" Property="Background" Value="{DynamicResource BgCardHover}"/>
                                <Setter TargetName="itemBorder" Property="BorderBrush" Value="{DynamicResource BorderInput}"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="itemBorder" Property="Background" Value="{DynamicResource BgCardHover}"/>
                                <Setter TargetName="itemBorder" Property="BorderBrush" Value="{DynamicResource AccentBlue}"/>
                                <Setter TargetName="itemBorder" Property="BorderThickness" Value="1.5"/>
                                <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Modern Flat DataGrid Styles -->
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="{DynamicResource BgCard}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderCard}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="RowBackground" Value="{DynamicResource DataGridRowBg}"/>
            <Setter Property="AlternatingRowBackground" Value="{DynamicResource DataGridAltRowBg}"/>
            <Setter Property="HorizontalGridLinesBrush" Value="{DynamicResource GridLines}"/>
            <Setter Property="VerticalGridLinesBrush" Value="Transparent"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="RowHeight" Value="34"/>
            <Setter Property="CanUserResizeRows" Value="False"/>
            <Setter Property="AutoGenerateColumns" Value="False"/>
        </Style>

        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="{DynamicResource DataGridHeaderBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource DataGridHeaderFg}"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Height" Value="34"/>
            <Setter Property="Padding" Value="10,0"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderCard}"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>

        <Style TargetType="DataGridRow">
            <Setter Property="SnapsToDevicePixels" Value="true"/>
            <Setter Property="Validation.ErrorTemplate" Value="{x:Null}"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="true">
                    <Setter Property="Background" Value="{DynamicResource BgCardHover}"/>
                    <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="DataGridCell">
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="10,4"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="DataGridCell">
                        <Border x:Name="cellBorder" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
                            <ContentPresenter VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="cellBorder" Property="Background" Value="{DynamicResource BgCardHover}"/>
                                <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid x:Name="mainGrid" Background="{DynamicResource BgApp}">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Top Header Bar -->
        <Border x:Name="topHeaderBorder" Grid.Row="0" Background="{DynamicResource BgHeader}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="0,0,0,1" Padding="18,12">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <!-- Brand & Utilities -->
                <Grid Grid.Row="0" Margin="0,0,0,12">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                        <Border Background="{DynamicResource AccentBlue}" CornerRadius="6" Width="30" Height="30" Margin="0,0,10,0">
                            <TextBlock Text="⇄" Foreground="#FFFFFF" FontSize="18" FontWeight="Bold" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <StackPanel VerticalAlignment="Center">
                            <TextBlock x:Name="txtAppTitle" Text="Excel Master Updater" Foreground="{DynamicResource TextPrimary}" FontSize="17" FontWeight="Bold"/>
                            <TextBlock x:Name="txtAppSubtitle" Text="Wielowątkowa synchronizacja i scalanie arkuszy danych" Foreground="{DynamicResource TextSecondary}" FontSize="11" Margin="0,1,0,0"/>
                        </StackPanel>
                        <Border x:Name="badgeProfile" Background="{DynamicResource BgCardHover}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="12" Padding="10,3" Margin="16,0,0,0" VerticalAlignment="Center" Visibility="Collapsed">
                            <TextBlock x:Name="txtProfileBadge" Text="Profil: Automatyczny" Foreground="{DynamicResource TextPrimary}" FontSize="11" FontWeight="Medium"/>
                        </Border>
                    </StackPanel>

                    <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                        <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Padding="8,4" Margin="0,0,10,0">
                            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                <TextBlock x:Name="lblLanguage" Text="Język:" Foreground="{DynamicResource TextSecondary}" VerticalAlignment="Center" Margin="0,0,6,0"/>
                                <ComboBox x:Name="cmbLanguage" Width="95" Background="{DynamicResource BgInput}" Foreground="{DynamicResource TextPrimary}" BorderThickness="0"/>
                            </StackPanel>
                        </Border>
                        <Button x:Name="btnThemeToggle" Content="🌙 Ciemny" Margin="0,0,6,0"/>
                        <Button x:Name="btnSettings" Content="⚙ Ustawienia" Margin="0,0,6,0"/>
                        <Button x:Name="btnRestoreBackup" Content="↺ Przywróć kopię" Margin="0,0,6,0"/>
                        <Button x:Name="btnOpenBackups" Content="📁 Kopie" Margin="0,0,6,0"/>
                        <Button x:Name="btnOpenLogs" Content="📋 Logi" Margin="0,0,6,0"/>
                        <Button x:Name="btnHelp" Content="❓ Pomoc"/>
                    </StackPanel>
                </Grid>

                <!-- Source File Cards (Side-by-Side) -->
                <Grid Grid.Row="1">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="12"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <!-- Card 1: Master Base File -->
                    <Border x:Name="cardBaseFile" Grid.Column="0" AllowDrop="True" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="12,10">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <Grid Grid.Row="0" Margin="0,0,0,8">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="140"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                    <TextBlock Text="📁" FontSize="13" Margin="0,0,6,0"/>
                                    <TextBlock x:Name="lblBasePath" Text="Baza główna" Foreground="{DynamicResource TextPrimary}" FontWeight="SemiBold"/>
                                </StackPanel>
                                <TextBlock x:Name="lblBaseSheet" Text="Arkusz:" Grid.Column="1" Foreground="{DynamicResource TextSecondary}" FontSize="12" VerticalAlignment="Center" Margin="0,0,6,0"/>
                                <ComboBox x:Name="cmbBaseSheet" Grid.Column="2" Background="{DynamicResource BgInput}" Foreground="{DynamicResource TextPrimary}" BorderBrush="{DynamicResource BorderInput}"/>
                            </Grid>
                            <Grid Grid.Row="1">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBox x:Name="txtBasePath" Grid.Column="0" Background="{DynamicResource BgInput}" Foreground="{DynamicResource TextPrimary}" BorderBrush="{DynamicResource BorderInput}" IsReadOnly="True" AllowDrop="True" Margin="0,0,8,0"/>
                                <Button x:Name="btnBrowseBase" Grid.Column="1" Content="Wybierz bazę..."/>
                            </Grid>
                        </Grid>
                    </Border>

                    <!-- Card 2: Incoming Update File -->
                    <Border x:Name="cardIncomingFile" Grid.Column="2" AllowDrop="True" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="12,10">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <Grid Grid.Row="0" Margin="0,0,0,8">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="140"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                    <TextBlock Text="📥" FontSize="13" Margin="0,0,6,0"/>
                                    <TextBlock x:Name="lblIncomingPath" Text="Plik zmian" Foreground="{DynamicResource TextPrimary}" FontWeight="SemiBold"/>
                                </StackPanel>
                                <TextBlock x:Name="lblIncomingSheet" Text="Arkusz:" Grid.Column="1" Foreground="{DynamicResource TextSecondary}" FontSize="12" VerticalAlignment="Center" Margin="0,0,6,0"/>
                                <ComboBox x:Name="cmbIncomingSheet" Grid.Column="2" Background="{DynamicResource BgInput}" Foreground="{DynamicResource TextPrimary}" BorderBrush="{DynamicResource BorderInput}"/>
                            </Grid>
                            <Grid Grid.Row="1">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBox x:Name="txtIncomingPath" Grid.Column="0" Background="{DynamicResource BgInput}" Foreground="{DynamicResource TextPrimary}" BorderBrush="{DynamicResource BorderInput}" AllowDrop="True" Margin="0,0,8,0"/>
                                <Button x:Name="btnBrowseIncoming" Grid.Column="1" Content="Wybierz zmiany..."/>
                            </Grid>
                        </Grid>
                    </Border>
                </Grid>
            </Grid>
        </Border>

        <!-- Main Tab Control -->
        <TabControl x:Name="mainTabs" Grid.Row="1">
            <!-- Tab 1: Column Mapping -->
            <TabItem x:Name="tabMapping" Header="1. Mapowanie kolumn">
                <Grid Margin="18">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="2*" MinHeight="90"/>
                        <RowDefinition Height="8"/>
                        <RowDefinition Height="3*" MinHeight="48"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <!-- Action Toolbar -->
                    <Border Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="12,8" Margin="0,0,0,10">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                <Button x:Name="btnAutoMap" Content="⚡ Automatyczne mapowanie" Background="{DynamicResource AccentBlue}" Foreground="#FFFFFF" FontWeight="SemiBold" BorderThickness="0" Margin="0,0,8,0"/>
                                <Button x:Name="btnAddRule" Content="+ Dodaj regułę" Margin="0,0,8,0"/>
                                <Button x:Name="btnEditRule" Content="Edytuj regułę" Margin="0,0,8,0"/>
                                <Button x:Name="btnRemoveRule" Content="Usuń regułę" Margin="0,0,8,0"/>
                                <Button x:Name="btnSaveProfile" Content="Zapisz profil"/>
                            </StackPanel>

                            <!-- Mapping Rules Search & Filter -->
                            <Border Grid.Column="1" Background="{DynamicResource BgInput}" BorderBrush="{DynamicResource BorderInput}" BorderThickness="1" CornerRadius="6" Padding="6,2" Width="220" VerticalAlignment="Center">
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Grid.Column="0" Text="🔍" FontSize="11" Foreground="{DynamicResource TextMuted}" VerticalAlignment="Center" Margin="2,0,5,0"/>
                                    <TextBox x:Name="txtSearchMapping" Grid.Column="1" Background="Transparent" Foreground="{DynamicResource TextPrimary}" BorderThickness="0" Padding="0,2" VerticalAlignment="Center"/>
                                    <Button x:Name="btnSearchMappingClear" Grid.Column="2" Content="✕" Width="16" Height="16" Padding="0" FontSize="9" Background="Transparent" Foreground="{DynamicResource TextMuted}" BorderThickness="0" Cursor="Hand" Visibility="Collapsed"/>
                                </Grid>
                            </Border>
                        </Grid>
                    </Border>

                    <!-- Mapping Grid & Join Key Split -->
                    <Grid Grid.Row="1">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="14"/>
                            <ColumnDefinition Width="340"/>
                        </Grid.ColumnDefinitions>

                        <!-- Mapping Rules DataGrid in Card -->
                        <Border Grid.Column="0" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="0">
                            <DataGrid x:Name="gridMappingRules" Margin="0" AutoGenerateColumns="False" CanUserAddRows="False"
                                      CanUserDeleteRows="False" CanUserReorderColumns="False" CanUserResizeRows="False"
                                      HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto">
                                <DataGrid.Columns>
                                    <DataGridTextColumn x:Name="colBase" Header="Kolumny w bazie" Binding="{Binding BaseColsStr}" Width="*"/>
                                    <DataGridTextColumn x:Name="colIncoming" Header="Kolumny w pliku zmian" Binding="{Binding UpdColsStr}" Width="*"/>
                                    <DataGridTextColumn x:Name="colMergeMode" Header="Tryb łączenia" Binding="{Binding MergeMode}" Width="120"/>
                                    <DataGridTextColumn x:Name="colSeparator" Header="Separator" Binding="{Binding Separator}" Width="80"/>
                                    <DataGridTextColumn x:Name="colDiffPolicy" Header="Zasada zmian" Binding="{Binding DiffPolicyDisplay}" Width="160"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </Border>

                        <!-- Join Key Picker in Card -->
                        <Border x:Name="joinBorder" Grid.Column="2" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="14">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>
                                <StackPanel Grid.Row="0" Margin="0,0,0,10">
                                    <TextBlock x:Name="lblJoinKeyTitle" Text="🔗 Klucz złączenia (Join Key)" Foreground="{DynamicResource TextPrimary}" FontWeight="Bold" FontSize="14"/>
                                    <TextBlock x:Name="lblJoinKeyHint" Text="Zaznacz kolumny jednoznacznie identyfikujące wiersz." Foreground="{DynamicResource TextSecondary}" FontSize="11" Margin="0,2,0,0"/>
                                </StackPanel>

                                <TextBlock x:Name="lblJoinBaseSub" Grid.Row="1" Text="Kolumny klucza bazy:" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="0,0,0,4"/>
                                <ListBox x:Name="lbJoinBase" Grid.Row="2" Background="{DynamicResource BgInput}" BorderBrush="{DynamicResource BorderInput}" SelectionMode="Extended" Margin="0,0,0,10"/>

                                <TextBlock x:Name="lblJoinIncomingSub" Grid.Row="3" Text="Kolumny klucza zmian:" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="0,0,0,4"/>
                                <ListBox x:Name="lbJoinIncoming" Grid.Row="4" Background="{DynamicResource BgInput}" BorderBrush="{DynamicResource BorderInput}" SelectionMode="Extended"/>
                            </Grid>
                        </Border>
                    </Grid>

                    <!-- Splitter between Mapping and Data Preview with Grab Handle -->
                    <GridSplitter Grid.Row="2" Height="8" HorizontalAlignment="Stretch" VerticalAlignment="Center"
                                  Background="Transparent" Cursor="SizeNS" ResizeDirection="Rows" ResizeBehavior="PreviousAndNext">
                        <GridSplitter.Template>
                            <ControlTemplate TargetType="GridSplitter">
                                <Border Background="Transparent" VerticalAlignment="Stretch" HorizontalAlignment="Stretch" Padding="0,2">
                                    <Grid VerticalAlignment="Center">
                                        <Rectangle x:Name="splitLine" Height="1" Fill="{DynamicResource BorderCard}" HorizontalAlignment="Stretch"/>
                                        <Border x:Name="splitPill" Width="48" Height="4" Background="{DynamicResource BorderCard}" CornerRadius="2" HorizontalAlignment="Center"/>
                                    </Grid>
                                </Border>
                                <ControlTemplate.Triggers>
                                    <Trigger Property="IsMouseOver" Value="True">
                                        <Setter TargetName="splitLine" Property="Fill" Value="{DynamicResource AccentBlue}"/>
                                        <Setter TargetName="splitPill" Property="Background" Value="{DynamicResource AccentBlue}"/>
                                    </Trigger>
                                </ControlTemplate.Triggers>
                            </ControlTemplate>
                        </GridSplitter.Template>
                    </GridSplitter>

                    <!-- Data & Mapping Live Preview Card -->
                    <Border x:Name="cardDataPreview" Grid.Row="3" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="12,10" Margin="0,4,0,0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>

                            <!-- Preview Header with Controls -->
                            <Grid Grid.Row="0" Margin="0,0,0,8">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>

                                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                    <TextBlock Text="📋 " FontSize="14" VerticalAlignment="Center"/>
                                    <TextBlock x:Name="txtPreviewHeader" Text="Podgląd danych i mapowania" Foreground="{DynamicResource TextPrimary}" FontWeight="Bold" FontSize="13" VerticalAlignment="Center"/>
                                    <Border Background="{DynamicResource BgCardHover}" CornerRadius="4" Padding="6,2" Margin="10,0,0,0" VerticalAlignment="Center">
                                        <TextBlock x:Name="txtPreviewHint" Text="Podgląd 1. wiersza bazy, 1. wiersza zmian oraz wynikowego mapowania na żywo" Foreground="{DynamicResource TextSecondary}" FontSize="11"/>
                                    </Border>
                                </StackPanel>

                                <TextBlock x:Name="txtPreviewInfo" Grid.Column="1" HorizontalAlignment="Center" VerticalAlignment="Center" Foreground="{DynamicResource TextSecondary}" FontSize="11"/>

                                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                                    <Button x:Name="btnPrevSampleRow" Content="◀ Poprz. wiersz" Margin="0,0,4,0" Padding="8,3" FontSize="11"/>
                                    <Button x:Name="btnNextSampleRow" Content="Nast. wiersz ▶" Margin="0,0,8,0" Padding="8,3" FontSize="11"/>
                                    <Button x:Name="btnRefreshPreview" Content="🔄 Odśwież podgląd" Background="{DynamicResource BgCardHover}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" Padding="10,3" FontSize="11"/>
                                </StackPanel>
                            </Grid>

                            <!-- TabControl for 3 Preview Modes -->
                            <TabControl x:Name="tabsMappingPreview" Grid.Row="1" Background="Transparent" BorderThickness="0">
                                <!-- Tab A: Mapping Result Projection -->
                                <TabItem x:Name="tabPrevResult" Header="⚡ Wynik mapowania (Podgląd)">
                                    <Grid Margin="0,6,0,0">
                                        <DataGrid x:Name="gridMappingResultPreview" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False"
                                                  CanUserReorderColumns="False" CanUserResizeRows="False" CanUserSortColumns="True"
                                                  HeadersVisibility="Column" GridLinesVisibility="All"
                                                  HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto"
                                                  SelectionMode="Single" SelectionUnit="FullRow"
                                                  Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1"
                                                  RowBackground="{DynamicResource DataGridRowBg}" AlternatingRowBackground="{DynamicResource DataGridRowBg}"
                                                  HorizontalGridLinesBrush="{DynamicResource GridLines}" VerticalGridLinesBrush="{DynamicResource GridLines}">
                                            <DataGrid.Columns>
                                                <DataGridTextColumn x:Name="colPrevTargetBase" Header="Kolumna docelowa (Baza)" Binding="{Binding TargetBaseColumn}" Width="180" FontWeight="SemiBold"/>
                                                <DataGridTextColumn x:Name="colPrevSourceExpr" Header="Zamapowane kolumny źródłowe" Binding="{Binding MappedSourceColumns}" Width="200"/>
                                                <DataGridTextColumn x:Name="colPrevProjectedVal" Header="Wartość wynikowa (z pliku zmian)" Binding="{Binding ProjectedValue}" Width="*"/>
                                                <DataGridTextColumn x:Name="colPrevCurrentBaseVal" Header="Dotychczasowa wartość w bazie" Binding="{Binding CurrentBaseValue}" Width="*"/>
                                                <DataGridTemplateColumn x:Name="colPrevStatus" Header="Status podglądu" Width="130">
                                                    <DataGridTemplateColumn.CellTemplate>
                                                        <DataTemplate>
                                                            <Border Background="{Binding StatusBg}" CornerRadius="4" Padding="6,2" HorizontalAlignment="Left" Margin="4,2">
                                                                <TextBlock Text="{Binding StatusText}" Foreground="{Binding StatusFg}" FontSize="11" FontWeight="Bold"/>
                                                            </Border>
                                                        </DataTemplate>
                                                    </DataGridTemplateColumn.CellTemplate>
                                                </DataGridTemplateColumn>
                                            </DataGrid.Columns>
                                        </DataGrid>
                                        <TextBlock x:Name="txtEmptyPreviewPrompt" Text="Wczytaj powyżej plik bazy oraz plik zmian, aby zobaczyć podgląd danych i wyniku mapowania."
                                                   Foreground="{DynamicResource TextSecondary}" FontSize="12" HorizontalAlignment="Center" VerticalAlignment="Center" Visibility="Collapsed"/>
                                    </Grid>
                                </TabItem>

                                <!-- Tab B: Incoming File (1st Row) -->
                                <TabItem x:Name="tabPrevIncoming" Header="📥 Plik zmian (Przykładowy wiersz)">
                                    <Grid Margin="0,6,0,0">
                                        <DataGrid x:Name="gridPrevIncomingRow" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False"
                                                  CanUserReorderColumns="False" CanUserResizeRows="False" CanUserSortColumns="False"
                                                  HeadersVisibility="Column" GridLinesVisibility="All"
                                                  HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto"
                                                  SelectionMode="Single" SelectionUnit="Cell"
                                                  Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1"
                                                  RowBackground="{DynamicResource DataGridRowBg}" AlternatingRowBackground="{DynamicResource DataGridRowBg}"
                                                  HorizontalGridLinesBrush="{DynamicResource GridLines}" VerticalGridLinesBrush="{DynamicResource GridLines}"/>
                                        <TextBlock x:Name="txtEmptyIncomingPrompt" Text="Wybierz plik zmian powyżej..." Foreground="{DynamicResource TextSecondary}" FontSize="12" HorizontalAlignment="Center" VerticalAlignment="Center" Visibility="Collapsed"/>
                                    </Grid>
                                </TabItem>

                                <!-- Tab C: Base File (1st Row) -->
                                <TabItem x:Name="tabPrevBase" Header="📁 Plik bazy (Przykładowy wiersz)">
                                    <Grid Margin="0,6,0,0">
                                        <DataGrid x:Name="gridPrevBaseRow" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False"
                                                  CanUserReorderColumns="False" CanUserResizeRows="False" CanUserSortColumns="False"
                                                  HeadersVisibility="Column" GridLinesVisibility="All"
                                                  HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto"
                                                  SelectionMode="Single" SelectionUnit="Cell"
                                                  Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1"
                                                  RowBackground="{DynamicResource DataGridRowBg}" AlternatingRowBackground="{DynamicResource DataGridRowBg}"
                                                  HorizontalGridLinesBrush="{DynamicResource GridLines}" VerticalGridLinesBrush="{DynamicResource GridLines}"/>
                                        <TextBlock x:Name="txtEmptyBasePrompt" Text="Wybierz plik bazy powyżej..." Foreground="{DynamicResource TextSecondary}" FontSize="12" HorizontalAlignment="Center" VerticalAlignment="Center" Visibility="Collapsed"/>
                                    </Grid>
                                </TabItem>
                            </TabControl>
                        </Grid>
                    </Border>

                    <Border Grid.Row="4" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="14,10" Margin="0,10,0,0">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Grid.Column="0" Orientation="Vertical" VerticalAlignment="Center" Margin="0,0,16,0">
                                <TextBlock x:Name="lblCompareOptionsTitle" Text="⚙️ Opcje trybu porównywania:" FontWeight="SemiBold" FontSize="12" Foreground="{DynamicResource TextSecondary}" Margin="0,0,0,6"/>
                                <WrapPanel Orientation="Horizontal" VerticalAlignment="Center">
                                    <CheckBox x:Name="chkIgnoreCase" Content="Ignoruj wielkość liter" IsChecked="True" Foreground="{DynamicResource TextPrimary}" Margin="0,0,18,2" VerticalContentAlignment="Center" Cursor="Hand"/>
                                    <CheckBox x:Name="chkTrimWhitespace" Content="Przycinaj białe znaki" IsChecked="True" Foreground="{DynamicResource TextPrimary}" Margin="0,0,18,2" VerticalContentAlignment="Center" Cursor="Hand"/>
                                    <CheckBox x:Name="chkIgnoreSpecialChars" Content="Ignoruj interpunkcję i znaki specjalne" IsChecked="True" Foreground="{DynamicResource TextPrimary}" Margin="0,0,18,2" VerticalContentAlignment="Center" Cursor="Hand"/>
                                    <CheckBox x:Name="chkIgnoreAllSpaces" Content="Ignoruj spacje wewnętrzne" IsChecked="True" Foreground="{DynamicResource TextPrimary}" Margin="0,0,18,2" VerticalContentAlignment="Center" Cursor="Hand"/>
                                </WrapPanel>
                            </StackPanel>
                            <Button x:Name="btnRunCompare" Grid.Column="1" Content="Rozpocznij porównanie ▶" Background="{DynamicResource AccentGreen}" Foreground="#FFFFFF" FontWeight="Bold" FontSize="14" BorderThickness="0" Padding="26,10" VerticalAlignment="Center" Cursor="Hand"/>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>

            <!-- Tab 2: Review & Approval -->
            <TabItem x:Name="tabReview" Header="2. Przegląd i zatwierdzanie">
                <Grid Margin="18">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <!-- Top KPI Counters & Filter Strip -->
                    <Border x:Name="filterBorder" Grid.Row="0" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="12,8" Margin="0,0,0,14">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <!-- KPI Interactive Stat Pills -->
                            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                <TextBlock x:Name="txtCounters" Visibility="Collapsed"/>
                                <Border x:Name="kpiPillAll" Background="{DynamicResource BgCardHover}" BorderBrush="{DynamicResource AccentBlue}" BorderThickness="2" CornerRadius="6" Padding="8,4" Margin="0,0,6,0" Cursor="Hand" ToolTip="Wszystkie rekordy">
                                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                        <TextBlock Text="📊" FontSize="11" Margin="0,0,4,0" VerticalAlignment="Center"/>
                                        <TextBlock x:Name="lblKpiAll" Text="Wszystkie:" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="0,0,4,0"/>
                                        <TextBlock x:Name="txtKpiCountAll" Text="0" Foreground="{DynamicResource TextPrimary}" FontWeight="Bold" FontSize="12"/>
                                    </StackPanel>
                                </Border>
                                <Border x:Name="kpiPillNew" Background="{DynamicResource BgCardHover}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Padding="8,4" Margin="0,0,6,0" Cursor="Hand" ToolTip="Tylko nowe rekordy">
                                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                        <TextBlock Text="✨" FontSize="11" Margin="0,0,4,0" VerticalAlignment="Center"/>
                                        <TextBlock x:Name="lblKpiNew" Text="Nowe:" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="0,0,4,0"/>
                                        <TextBlock x:Name="txtKpiCountNew" Text="0" Foreground="{DynamicResource AccentGreen}" FontWeight="Bold" FontSize="12"/>
                                    </StackPanel>
                                </Border>
                                <Border x:Name="kpiPillChanged" Background="{DynamicResource BgCardHover}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Padding="8,4" Margin="0,0,6,0" Cursor="Hand" ToolTip="Tylko zmienione rekordy">
                                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                        <TextBlock Text="⚡" FontSize="11" Margin="0,0,4,0" VerticalAlignment="Center"/>
                                        <TextBlock x:Name="lblKpiChanged" Text="Zmienione:" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="0,0,4,0"/>
                                        <TextBlock x:Name="txtKpiCountChanged" Text="0" Foreground="{DynamicResource AccentBlue}" FontWeight="Bold" FontSize="12"/>
                                    </StackPanel>
                                </Border>
                                <Border x:Name="kpiPillAmbiguous" Background="{DynamicResource BgCardHover}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Padding="8,4" Margin="0,0,6,0" Cursor="Hand" ToolTip="Niejednoznaczne dopasowania">
                                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                        <TextBlock Text="❓" FontSize="11" Margin="0,0,4,0" VerticalAlignment="Center"/>
                                        <TextBlock x:Name="lblKpiAmbiguous" Text="Niejednoznaczne:" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="0,0,4,0"/>
                                        <TextBlock x:Name="txtKpiCountAmbiguous" Text="0" Foreground="{DynamicResource AccentAmber}" FontWeight="Bold" FontSize="12"/>
                                    </StackPanel>
                                </Border>
                                <Border x:Name="kpiPillAccepted" Background="{DynamicResource BgCardHover}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Padding="8,4" Margin="0,0,6,0" Cursor="Hand" ToolTip="Zaakceptowane zmiany">
                                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                        <TextBlock Text="✔" FontSize="11" Margin="0,0,4,0" VerticalAlignment="Center"/>
                                        <TextBlock x:Name="lblKpiAccepted" Text="Zaakceptowane:" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="0,0,4,0"/>
                                        <TextBlock x:Name="txtKpiCountAccepted" Text="0" Foreground="{DynamicResource AccentGreen}" FontWeight="Bold" FontSize="12"/>
                                    </StackPanel>
                                </Border>
                                <Border x:Name="kpiPillSkipped" Background="{DynamicResource BgCardHover}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Padding="8,4" Margin="0,0,6,0" Cursor="Hand" ToolTip="Pominięte lub odrzucone">
                                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                        <TextBlock Text="⏭" FontSize="11" Margin="0,0,4,0" VerticalAlignment="Center"/>
                                        <TextBlock x:Name="lblKpiSkipped" Text="Pominięte:" Foreground="{DynamicResource TextSecondary}" FontSize="12" Margin="0,0,4,0"/>
                                        <TextBlock x:Name="txtKpiCountSkipped" Text="0" Foreground="{DynamicResource TextMuted}" FontWeight="Bold" FontSize="12"/>
                                    </StackPanel>
                                </Border>
                            </StackPanel>

                            <!-- Search & Filter Controls -->
                            <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                                <Border Background="{DynamicResource BgInput}" BorderBrush="{DynamicResource BorderInput}" BorderThickness="1" CornerRadius="6" Padding="6,2" Margin="0,0,10,0" Width="220">
                                    <Grid>
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="Auto"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBlock Grid.Column="0" Text="🔍" FontSize="11" Foreground="{DynamicResource TextMuted}" VerticalAlignment="Center" Margin="2,0,5,0"/>
                                        <TextBox x:Name="txtSearchReview" Grid.Column="1" Background="Transparent" Foreground="{DynamicResource TextPrimary}" BorderThickness="0" Padding="0,2" VerticalAlignment="Center"/>
                                        <Button x:Name="btnSearchClear" Grid.Column="2" Content="✕" Width="16" Height="16" Padding="0" FontSize="9" Background="Transparent" Foreground="{DynamicResource TextMuted}" BorderThickness="0" Cursor="Hand" Visibility="Collapsed"/>
                                    </Grid>
                                </Border>
                                <ComboBox x:Name="cmbFilterStatus" Width="150" SelectedIndex="0">
                                    <ComboBoxItem x:Name="cbiFilterAll" Content="Wszystkie" Tag="All"/>
                                    <ComboBoxItem x:Name="cbiFilterNew" Content="Nowe" Tag="New"/>
                                    <ComboBoxItem x:Name="cbiFilterChanged" Content="Zmienione" Tag="Changed"/>
                                    <ComboBoxItem x:Name="cbiFilterRemoved" Content="Usunięte" Tag="Removed"/>
                                    <ComboBoxItem x:Name="cbiFilterAmbiguous" Content="Niejednoznaczne" Tag="Ambiguous"/>
                                    <ComboBoxItem x:Name="cbiFilterUnchanged" Content="Bez zmian" Tag="Unchanged"/>
                                    <ComboBoxItem x:Name="cbiFilterAccepted" Content="Zaakceptowane" Tag="Accepted"/>
                                    <ComboBoxItem x:Name="cbiFilterSkipped" Content="Pominięte" Tag="Skipped"/>
                                    <ComboBoxItem x:Name="cbiFilterRejected" Content="Odrzucone" Tag="Rejected"/>
                                </ComboBox>
                                <CheckBox x:Name="chkShowUnchanged" Margin="10,0,0,0" VerticalAlignment="Center" Content="Pokaż bez zmian" Foreground="{DynamicResource TextPrimary}"/>
                            </StackPanel>
                        </Grid>
                    </Border>

                    <!-- Review Master-Detail -->
                    <Grid Grid.Row="1">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="380"/>
                            <ColumnDefinition Width="14"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>

                        <!-- Items List -->
                        <Border Grid.Column="0" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="6">
                            <ListBox x:Name="lbReviewItems" BorderThickness="0" Background="Transparent">
                                <ListBox.ItemTemplate>
                                    <DataTemplate>
                                        <Grid>
                                            <Grid.ColumnDefinitions>
                                                <ColumnDefinition Width="36"/>
                                                <ColumnDefinition Width="*"/>
                                                <ColumnDefinition Width="Auto"/>
                                            </Grid.ColumnDefinitions>
                                            <Border Background="{DynamicResource BgCardHover}" CornerRadius="4" Width="28" Height="24" HorizontalAlignment="Left" VerticalAlignment="Center">
                                                <TextBlock Text="{Binding IndexStr}" Foreground="{DynamicResource TextMuted}" FontSize="11" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                            </Border>
                                            <StackPanel Grid.Column="1" VerticalAlignment="Center" Margin="6,0,8,0">
                                                <TextBlock Text="{Binding Title}" Foreground="{DynamicResource TextPrimary}" FontWeight="SemiBold" FontSize="13" TextTrimming="CharacterEllipsis"/>
                                                <TextBlock Text="{Binding Subtitle}" Foreground="{DynamicResource TextSecondary}" FontSize="11" Margin="0,1,0,0" TextTrimming="CharacterEllipsis"/>
                                            </StackPanel>
                                            <Border Grid.Column="2" Background="{Binding StatusBg}" CornerRadius="4" Padding="7,3" VerticalAlignment="Center">
                                                <TextBlock Text="{Binding StatusText}" Foreground="{Binding StatusFg}" FontSize="10" FontWeight="Bold"/>
                                            </Border>
                                        </Grid>
                                    </DataTemplate>
                                </ListBox.ItemTemplate>
                            </ListBox>
                        </Border>

                        <!-- Right Detail Pane Card -->
                        <Border x:Name="detailBorder" Grid.Column="2" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="18">
                            <ScrollViewer VerticalScrollBarVisibility="Auto">
                                <StackPanel x:Name="panelDetailContent">
                                    <!-- Batch Column Toggles Banner -->
                                    <Border x:Name="cardBatchToggles" Background="{DynamicResource BgCardHover}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="6" Padding="10,8" Margin="0,0,0,14" Visibility="Collapsed">
                                        <StackPanel>
                                            <TextBlock x:Name="lblBatchColToggles" Text="Szybkie przełączanie kolumn (przejmij wybrane pola):" Foreground="{DynamicResource TextSecondary}" FontSize="11" FontWeight="SemiBold" Margin="0,0,0,6"/>
                                            <WrapPanel x:Name="wrapBatchColToggles" Orientation="Horizontal"/>
                                        </StackPanel>
                                    </Border>
                                    <Border BorderBrush="{DynamicResource BorderCard}" BorderThickness="0,0,0,1" Padding="0,0,0,12" Margin="0,0,0,12">
                                        <TextBlock x:Name="txtDetailHeader" Text="Wybierz wiersz z listy po lewej stronie" Foreground="{DynamicResource TextPrimary}" FontSize="16" FontWeight="Bold"/>
                                    </Border>
                                    <StackPanel x:Name="panelDiffContainer"/>
                                </StackPanel>
                            </ScrollViewer>
                        </Border>
                    </Grid>

                    <!-- Whole Base Row Preview & Direct Cell Editor -->
                    <Border x:Name="baseRowBorder" Grid.Row="2" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="12,10" Margin="0,10,0,0" Visibility="Collapsed">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <!-- Header -->
                            <Grid Grid.Row="0" Margin="0,0,0,8">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                    <TextBlock Text="📋 " FontSize="14" VerticalAlignment="Center"/>
                                    <TextBlock x:Name="txtBaseRowHeader" Text="Pełny wiersz z bazy danych" Foreground="{DynamicResource TextPrimary}" FontWeight="Bold" FontSize="13" VerticalAlignment="Center"/>
                                    <Border Background="{DynamicResource BgCardHover}" CornerRadius="4" Padding="6,2" Margin="10,0,0,0" VerticalAlignment="Center">
                                        <TextBlock x:Name="txtBaseRowHint" Text="Kliknij dwukrotnie komórkę lub użyj przycisku, aby zmodyfikować dowolne pole bazy" Foreground="{DynamicResource TextSecondary}" FontSize="11"/>
                                    </Border>
                                </StackPanel>
                                <StackPanel Grid.Column="2" Orientation="Horizontal">
                                    <Button x:Name="btnModifyBaseField" Content="✏ Modyfikuj pole w bazie..." Background="{DynamicResource AccentAmber}" Foreground="#FFFFFF" FontWeight="SemiBold" BorderThickness="0" Padding="12,4"/>
                                </StackPanel>
                            </Grid>
                            <!-- DataGrid showing the entire base row with all columns -->
                            <DataGrid x:Name="dgBaseFullRow" Grid.Row="1" Height="86" MaxHeight="110"
                                      AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False"
                                      CanUserReorderColumns="False" CanUserResizeRows="False" CanUserSortColumns="False"
                                      HeadersVisibility="Column"
                                      GridLinesVisibility="All" HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto"
                                      SelectionMode="Single" SelectionUnit="Cell"
                                      Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1"
                                      RowBackground="{DynamicResource DataGridRowBg}" AlternatingRowBackground="{DynamicResource DataGridRowBg}"
                                      HorizontalGridLinesBrush="{DynamicResource GridLines}" VerticalGridLinesBrush="{DynamicResource GridLines}"/>
                        </Grid>
                    </Border>

                    <!-- Bottom Action Controls -->
                    <Border x:Name="actionBorder" Grid.Row="3" Background="{DynamicResource BgCard}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="1" CornerRadius="8" Padding="12,10" Margin="0,10,0,0">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Orientation="Horizontal">
                                <Button x:Name="btnBackRow" Content="◀ Wstecz (B)" Margin="0,0,6,0"/>
                                <Button x:Name="btnAcceptRow" Content="✔ Akceptuj (A)" Background="{DynamicResource AccentGreen}" Foreground="#FFFFFF" FontWeight="Bold" BorderThickness="0" Margin="0,0,6,0"/>
                                <Button x:Name="btnRejectRow" Content="✖ Odrzuć (R)" Background="{DynamicResource AccentRed}" Foreground="#FFFFFF" FontWeight="Bold" BorderThickness="0" Margin="0,0,6,0"/>
                                <Button x:Name="btnSkipRow" Content="⏭ Pomiń (S)" Margin="0,0,6,0"/>
                                <Button x:Name="btnEditRow" Content="✏ Zmień wartość... (E)" Background="{DynamicResource AccentAmber}" Foreground="#FFFFFF" BorderThickness="0" Margin="0,0,6,0"/>
                                <Button x:Name="btnUndo" Content="↩ Cofnij (Ctrl+Z)" IsEnabled="False"/>
                            </StackPanel>
                            <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                                    <Button x:Name="btnAcceptAll" Content="✔✔ Akceptuj wszystkie" Background="{DynamicResource AccentGreen}" Foreground="#FFFFFF" BorderThickness="0" Margin="0,0,6,0"/>
                                    <Button x:Name="btnRejectAll" Content="✖✖ Odrzuć wszystkie" Background="{DynamicResource AccentRed}" Foreground="#FFFFFF" BorderThickness="0" Margin="0,0,6,0"/>
                                    <Button x:Name="btnExportReport" Content="📊 Eksportuj raport..."/>
                                </StackPanel>
                                <TextBlock x:Name="txtStagingSummary" Text="" HorizontalAlignment="Center" Foreground="{DynamicResource TextSecondary}" FontSize="11" Margin="0,4,0,0"/>
                            </StackPanel>
                            <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                                <Button x:Name="btnApplyAccepted" Content="Zastosuj do bazy" Background="{DynamicResource AccentBlue}" Foreground="#FFFFFF" FontWeight="Bold" FontSize="13" BorderThickness="0" Padding="14,8" Margin="0,0,6,0" IsEnabled="False"/>
                                <Button x:Name="btnApplyToNewFile" Content="💾 Zapisz do nowego pliku..." Background="{DynamicResource AccentPurple}" Foreground="#FFFFFF" FontWeight="Bold" FontSize="13" BorderThickness="0" Padding="14,8" IsEnabled="False"/>
                            </StackPanel>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>
        </TabControl>

        <!-- Status Bar -->
        <StatusBar x:Name="statusBar" Grid.Row="2" Background="{DynamicResource StatusBarBg}" Foreground="{DynamicResource StatusBarFg}" BorderBrush="{DynamicResource BorderCard}" BorderThickness="0,1,0,0" Padding="14,6">
            <StatusBarItem>
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="●" Foreground="{DynamicResource AccentGreen}" FontSize="10" Margin="0,0,8,0" VerticalAlignment="Center"/>
                    <TextBlock x:Name="txtStatusMsg" Text="Gotowy do pracy." Foreground="{DynamicResource StatusBarFg}"/>
                </StackPanel>
            </StatusBarItem>
        </StatusBar>
    </Grid>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [System.Windows.Markup.XamlReader]::Load($reader)
    $script:ActiveWindow = $window

    # Initialize State
    $script:AppConfig = $cfg
    $script:CurrentTheme = if ($cfg.Theme) { $cfg.Theme } else { 'Dark' }

    # Apply Dwm dark/light mode
    $hwnd = (New-Object System.Windows.Interop.WindowInteropHelper($window)).EnsureHandle()
    Set-WindowDwmTheme -Hwnd $hwnd -IsDark ($script:CurrentTheme -eq 'Dark')
    $script:AllReviewItems = [System.Collections.Generic.List[object]]::new()
    $script:BaseHeaders = @()
    $script:BaseDataRows = [System.Collections.Generic.List[object]]::new()
    $script:IncomingHeaders = @()
    $script:IncomingDataRows = [System.Collections.Generic.List[object]]::new()
    $script:CurrentProfile = $null
    $script:MappingRules = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
    # P3 — In-session undo stack: each entry is [ItemRef, PrevDecision, PrevStatusText, PrevStatusBg, PrevStatusFg]
    $script:UndoStack = [System.Collections.Generic.Stack[object]]::new()

    $script:window      = $window
    $mainGrid           = $window.FindName('mainGrid')
    $topHeaderBorder    = $window.FindName('topHeaderBorder')
    $cardBaseFile       = $window.FindName('cardBaseFile')
    $cardIncomingFile   = $window.FindName('cardIncomingFile')
        $txtAppTitle        = $window.FindName('txtAppTitle')
    $txtAppSubtitle     = $window.FindName('txtAppSubtitle')
    $lblBasePath        = $window.FindName('lblBasePath')
    $script:txtBasePath = $window.FindName('txtBasePath')
    $txtBasePath        = $script:txtBasePath
    $lblBaseSheet       = $window.FindName('lblBaseSheet')
    $script:cmbBaseSheet = $window.FindName('cmbBaseSheet')
    $cmbBaseSheet       = $script:cmbBaseSheet
    $btnBrowseBase      = $window.FindName('btnBrowseBase')
    $lblIncomingPath    = $window.FindName('lblIncomingPath')
    $script:txtIncomingPath = $window.FindName('txtIncomingPath')
    $txtIncomingPath    = $script:txtIncomingPath
    $lblIncomingSheet   = $window.FindName('lblIncomingSheet')
    $script:cmbIncomingSheet = $window.FindName('cmbIncomingSheet')
    $cmbIncomingSheet   = $script:cmbIncomingSheet
    $btnBrowseIncoming  = $window.FindName('btnBrowseIncoming')
    $badgeProfile       = $window.FindName('badgeProfile')
    $txtProfileBadge    = $window.FindName('txtProfileBadge')
    $lblLanguage        = $window.FindName('lblLanguage')
    $cmbLanguage        = $window.FindName('cmbLanguage')
    $btnThemeToggle     = $window.FindName('btnThemeToggle')
    $btnSettings        = $window.FindName('btnSettings')
    $btnRestoreBackup   = $window.FindName('btnRestoreBackup')
    $btnOpenBackups     = $window.FindName('btnOpenBackups')
    $btnOpenLogs        = $window.FindName('btnOpenLogs')
    $btnHelp            = $window.FindName('btnHelp')
    $script:mainTabs    = $window.FindName('mainTabs')
    $mainTabs           = $script:mainTabs
    $tabMapping         = $window.FindName('tabMapping')
    $tabReview          = $window.FindName('tabReview')
    $btnAutoMap         = $window.FindName('btnAutoMap')
    $btnAddRule         = $window.FindName('btnAddRule')
    $btnEditRule        = $window.FindName('btnEditRule')
    $btnRemoveRule      = $window.FindName('btnRemoveRule')
    $btnSaveProfile     = $window.FindName('btnSaveProfile')
    $script:txtSearchMapping = $window.FindName('txtSearchMapping')
    $txtSearchMapping        = $script:txtSearchMapping
    $script:btnSearchMappingClear = $window.FindName('btnSearchMappingClear')
    $btnSearchMappingClear   = $script:btnSearchMappingClear
    $gridMappingRules   = $window.FindName('gridMappingRules')
    $colBase            = $window.FindName('colBase')
    $colIncoming        = $window.FindName('colIncoming')
    $colMergeMode       = $window.FindName('colMergeMode')
    $colSeparator       = $window.FindName('colSeparator')
    $colDiffPolicy      = $window.FindName('colDiffPolicy')
    $joinBorder         = $window.FindName('joinBorder')
        $lblJoinKeyTitle    = $window.FindName('lblJoinKeyTitle')
    $lblJoinKeyHint     = $window.FindName('lblJoinKeyHint')
    $lblJoinBaseSub     = $window.FindName('lblJoinBaseSub')
    $script:lbJoinBase  = $window.FindName('lbJoinBase')
    $lbJoinBase         = $script:lbJoinBase
    $lblJoinIncomingSub = $window.FindName('lblJoinIncomingSub')
    $script:lbJoinIncoming = $window.FindName('lbJoinIncoming')
    $lbJoinIncoming     = $script:lbJoinIncoming
    $lblCompareOptionsTitle       = $window.FindName('lblCompareOptionsTitle')
    $script:chkIgnoreCase         = $window.FindName('chkIgnoreCase')
    $chkIgnoreCase                = $script:chkIgnoreCase
    $script:chkTrimWhitespace     = $window.FindName('chkTrimWhitespace')
    $chkTrimWhitespace            = $script:chkTrimWhitespace
    $script:chkIgnoreSpecialChars = $window.FindName('chkIgnoreSpecialChars')
    $chkIgnoreSpecialChars        = $script:chkIgnoreSpecialChars
    $script:chkIgnoreAllSpaces    = $window.FindName('chkIgnoreAllSpaces')
    $chkIgnoreAllSpaces           = $script:chkIgnoreAllSpaces
    if ($chkIgnoreCase -and $null -ne $script:AppConfig.CompareIgnoreCase) { $chkIgnoreCase.IsChecked = [bool]$script:AppConfig.CompareIgnoreCase }
    if ($chkTrimWhitespace -and $null -ne $script:AppConfig.CompareTrimWhitespace) { $chkTrimWhitespace.IsChecked = [bool]$script:AppConfig.CompareTrimWhitespace }
    if ($chkIgnoreSpecialChars -and $null -ne $script:AppConfig.CompareIgnoreSpecialChars) { $chkIgnoreSpecialChars.IsChecked = [bool]$script:AppConfig.CompareIgnoreSpecialChars }
    if ($chkIgnoreAllSpaces -and $null -ne $script:AppConfig.CompareIgnoreAllSpaces) { $chkIgnoreAllSpaces.IsChecked = [bool]$script:AppConfig.CompareIgnoreAllSpaces }
    $btnRunCompare      = $window.FindName('btnRunCompare')
    $filterBorder       = $window.FindName('filterBorder')
    $script:txtCounters = $window.FindName('txtCounters')
    $txtCounters        = $script:txtCounters
    $kpiPillAll         = $window.FindName('kpiPillAll')
    $lblKpiAll          = $window.FindName('lblKpiAll')
    $txtKpiCountAll     = $window.FindName('txtKpiCountAll')
    $kpiPillNew         = $window.FindName('kpiPillNew')
    $lblKpiNew          = $window.FindName('lblKpiNew')
    $txtKpiCountNew     = $window.FindName('txtKpiCountNew')
    $kpiPillChanged     = $window.FindName('kpiPillChanged')
    $lblKpiChanged      = $window.FindName('lblKpiChanged')
    $txtKpiCountChanged = $window.FindName('txtKpiCountChanged')
    $kpiPillAmbiguous   = $window.FindName('kpiPillAmbiguous')
    $lblKpiAmbiguous    = $window.FindName('lblKpiAmbiguous')
    $txtKpiCountAmbiguous = $window.FindName('txtKpiCountAmbiguous')
    $kpiPillAccepted    = $window.FindName('kpiPillAccepted')
    $lblKpiAccepted     = $window.FindName('lblKpiAccepted')
    $txtKpiCountAccepted = $window.FindName('txtKpiCountAccepted')
    $kpiPillSkipped     = $window.FindName('kpiPillSkipped')
    $lblKpiSkipped      = $window.FindName('lblKpiSkipped')
    $txtKpiCountSkipped = $window.FindName('txtKpiCountSkipped')
    $txtSearchReview    = $window.FindName('txtSearchReview')
    $btnSearchClear     = $window.FindName('btnSearchClear')
    $script:cmbFilterStatus = $window.FindName('cmbFilterStatus')
    $cmbFilterStatus    = $script:cmbFilterStatus
    $cbiFilterAll       = $window.FindName('cbiFilterAll')
    $cbiFilterNew       = $window.FindName('cbiFilterNew')
    $cbiFilterChanged   = $window.FindName('cbiFilterChanged')
    $cbiFilterRemoved   = $window.FindName('cbiFilterRemoved')
    $cbiFilterAmbiguous = $window.FindName('cbiFilterAmbiguous')
    $cbiFilterUnchanged = $window.FindName('cbiFilterUnchanged')
    $cbiFilterAccepted  = $window.FindName('cbiFilterAccepted')
    $cbiFilterSkipped   = $window.FindName('cbiFilterSkipped')
    $cbiFilterRejected  = $window.FindName('cbiFilterRejected')
    $script:chkShowUnchanged = $window.FindName('chkShowUnchanged')
    $chkShowUnchanged   = $script:chkShowUnchanged
    if ($chkShowUnchanged) {
        $chkShowUnchanged.IsChecked = ($script:AppConfig.ShowUnchangedRows -eq $true -or $script:AppConfig.AutoSkipUnchanged -eq $false)
    }
    $script:lbReviewItems = $window.FindName('lbReviewItems')
    $lbReviewItems      = $script:lbReviewItems
    $detailBorder       = $window.FindName('detailBorder')
    $cardBatchToggles   = $window.FindName('cardBatchToggles')
    $lblBatchColToggles = $window.FindName('lblBatchColToggles')
    $wrapBatchColToggles = $window.FindName('wrapBatchColToggles')
    $txtDetailHeader    = $window.FindName('txtDetailHeader')
    $panelDiffContainer = $window.FindName('panelDiffContainer')
    $baseRowBorder      = $window.FindName('baseRowBorder')
    $txtBaseRowHeader   = $window.FindName('txtBaseRowHeader')
    $txtBaseRowHint     = $window.FindName('txtBaseRowHint')
    $btnModifyBaseField = $window.FindName('btnModifyBaseField')
    $dgBaseFullRow      = $window.FindName('dgBaseFullRow')
    $actionBorder       = $window.FindName('actionBorder')
    $btnBackRow         = $window.FindName('btnBackRow')
    $btnAcceptRow       = $window.FindName('btnAcceptRow')
    $btnRejectRow       = $window.FindName('btnRejectRow')
    $btnSkipRow         = $window.FindName('btnSkipRow')
    $btnEditRow         = $window.FindName('btnEditRow')
    $btnUndo            = $window.FindName('btnUndo')
    $btnAcceptAll       = $window.FindName('btnAcceptAll')
    $btnRejectAll       = $window.FindName('btnRejectAll')
    $btnExportReport    = $window.FindName('btnExportReport')
    $txtStagingSummary  = $window.FindName('txtStagingSummary')
    $btnApplyAccepted   = $window.FindName('btnApplyAccepted')
    $btnApplyToNewFile  = $window.FindName('btnApplyToNewFile')
    $statusBar          = $window.FindName('statusBar')
    $script:txtStatusMsg = $window.FindName('txtStatusMsg')
    $txtStatusMsg       = $script:txtStatusMsg

    # Preview controls
    $script:cardDataPreview          = $window.FindName('cardDataPreview')
    $cardDataPreview                 = $script:cardDataPreview
    $script:txtPreviewHeader         = $window.FindName('txtPreviewHeader')
    $txtPreviewHeader                = $script:txtPreviewHeader
    $script:txtPreviewHint           = $window.FindName('txtPreviewHint')
    $txtPreviewHint                  = $script:txtPreviewHint
    $script:txtPreviewInfo           = $window.FindName('txtPreviewInfo')
    $txtPreviewInfo                  = $script:txtPreviewInfo
    $script:btnPrevSampleRow         = $window.FindName('btnPrevSampleRow')
    $btnPrevSampleRow                = $script:btnPrevSampleRow
    $script:btnNextSampleRow         = $window.FindName('btnNextSampleRow')
    $btnNextSampleRow                = $script:btnNextSampleRow
    $script:btnRefreshPreview        = $window.FindName('btnRefreshPreview')
    $btnRefreshPreview               = $script:btnRefreshPreview
    $script:tabsMappingPreview       = $window.FindName('tabsMappingPreview')
    $tabsMappingPreview              = $script:tabsMappingPreview
    $script:tabPrevResult            = $window.FindName('tabPrevResult')
    $tabPrevResult                   = $script:tabPrevResult
    $script:tabPrevIncoming          = $window.FindName('tabPrevIncoming')
    $tabPrevIncoming                 = $script:tabPrevIncoming
    $script:tabPrevBase              = $window.FindName('tabPrevBase')
    $tabPrevBase                     = $script:tabPrevBase
    $script:gridMappingResultPreview = $window.FindName('gridMappingResultPreview')
    $gridMappingResultPreview        = $script:gridMappingResultPreview
    $script:gridPrevIncomingRow      = $window.FindName('gridPrevIncomingRow')
    $gridPrevIncomingRow             = $script:gridPrevIncomingRow
    $script:gridPrevBaseRow          = $window.FindName('gridPrevBaseRow')
    $gridPrevBaseRow                 = $script:gridPrevBaseRow
    $script:txtEmptyPreviewPrompt    = $window.FindName('txtEmptyPreviewPrompt')
    $txtEmptyPreviewPrompt           = $script:txtEmptyPreviewPrompt
    $script:txtEmptyIncomingPrompt   = $window.FindName('txtEmptyIncomingPrompt')
    $txtEmptyIncomingPrompt          = $script:txtEmptyIncomingPrompt
    $script:txtEmptyBasePrompt       = $window.FindName('txtEmptyBasePrompt')
    $txtEmptyBasePrompt              = $script:txtEmptyBasePrompt
    $script:colPrevTargetBase        = $window.FindName('colPrevTargetBase')
    $colPrevTargetBase               = $script:colPrevTargetBase
    $script:colPrevSourceExpr        = $window.FindName('colPrevSourceExpr')
    $colPrevSourceExpr               = $script:colPrevSourceExpr
    $script:colPrevProjectedVal      = $window.FindName('colPrevProjectedVal')
    $colPrevProjectedVal             = $script:colPrevProjectedVal
    $script:colPrevCurrentBaseVal    = $window.FindName('colPrevCurrentBaseVal')
    $colPrevCurrentBaseVal           = $script:colPrevCurrentBaseVal
    $script:colPrevStatus            = $window.FindName('colPrevStatus')
    $colPrevStatus                   = $script:colPrevStatus

    # Bind Grid source
    $gridMappingRules.ItemsSource = $script:MappingRules

    # Helper: Apply theme colors
    $ApplyTheme = {
        param([string]$ThemeName)
        $script:CurrentTheme = $ThemeName
        $p = Get-UpdaterThemePalette $ThemeName
        $isDark = $p.IsDark
        $w = if ($script:ActiveWindow) { $script:ActiveWindow } elseif ($script:window) { $script:window } else { $window }
        if (-not $w) { return }

        # Update dynamic brushes in Window Resources (freeze each brush for thread-safety and performance)
        foreach ($k in $p.Keys) {
            if ($k -ne 'IsDark') {
                $c = [System.Windows.Media.ColorConverter]::ConvertFromString($p[$k])
                $brush = [System.Windows.Media.SolidColorBrush]::new($c)
                $brush.Freeze()
                $w.Resources[$k] = $brush
            }
        }

        $setBrush = {
            param($elem, $prop, $resKey)
            if (-not $elem) { return }
            $b = $w.Resources[$resKey]
            if (-not $b) { return }
            $ctrl = if ($elem -is [string]) { $w.FindName($elem) } else { $elem }
            if ($ctrl) {
                try { $ctrl.$prop = $b } catch {}
            }
        }

        # Apply root canvas backgrounds & foregrounds
        try {
            $w.Background = $w.Resources['BgApp']
            $w.Foreground = $w.Resources['TextPrimary']
        } catch {}

        & $setBrush 'mainGrid' 'Background' 'BgApp'
        & $setBrush 'topHeaderBorder' 'Background' 'BgHeader'
        & $setBrush 'topHeaderBorder' 'BorderBrush' 'BorderCard'
        & $setBrush 'mainTabs' 'Background' 'BgApp'
        & $setBrush 'joinBorder' 'Background' 'BgCard'
        & $setBrush 'joinBorder' 'BorderBrush' 'BorderCard'
        & $setBrush 'filterBorder' 'Background' 'BgCard'
        & $setBrush 'filterBorder' 'BorderBrush' 'BorderCard'
        & $setBrush 'detailBorder' 'Background' 'BgCard'
        & $setBrush 'detailBorder' 'BorderBrush' 'BorderCard'
        & $setBrush 'baseRowBorder' 'Background' 'BgCard'
        & $setBrush 'baseRowBorder' 'BorderBrush' 'BorderCard'
        & $setBrush 'actionBorder' 'Background' 'BgCard'
        & $setBrush 'actionBorder' 'BorderBrush' 'BorderCard'
        & $setBrush 'statusBar' 'Background' 'StatusBarBg'
        & $setBrush 'statusBar' 'Foreground' 'StatusBarFg'
        & $setBrush 'statusBar' 'BorderBrush' 'BorderCard'

        # Text labels and titles
        & $setBrush 'txtAppTitle' 'Foreground' 'TextPrimary'
        & $setBrush 'lblBasePath' 'Foreground' 'TextSecondary'
        & $setBrush 'lblBaseSheet' 'Foreground' 'TextSecondary'
        & $setBrush 'lblIncomingPath' 'Foreground' 'TextSecondary'
        & $setBrush 'lblIncomingSheet' 'Foreground' 'TextSecondary'
        & $setBrush 'lblLanguage' 'Foreground' 'TextSecondary'
        & $setBrush 'lblJoinKeyTitle' 'Foreground' 'TextPrimary'
        & $setBrush 'lblJoinBaseSub' 'Foreground' 'TextSecondary'
        & $setBrush 'lblJoinIncomingSub' 'Foreground' 'TextSecondary'
        & $setBrush 'txtCounters' 'Foreground' 'TextPrimary'
        & $setBrush 'txtDetailHeader' 'Foreground' 'TextPrimary'
        & $setBrush 'txtStatusMsg' 'Foreground' 'StatusBarFg'
        & $setBrush 'tabMapping' 'Foreground' 'TextPrimary'
        & $setBrush 'tabReview' 'Foreground' 'TextPrimary'

        # Comparison mode controls
        & $setBrush 'lblCompareOptionsTitle' 'Foreground' 'TextSecondary'
        & $setBrush 'chkIgnoreCase' 'Foreground' 'TextPrimary'
        & $setBrush 'chkTrimWhitespace' 'Foreground' 'TextPrimary'
        & $setBrush 'chkIgnoreSpecialChars' 'Foreground' 'TextPrimary'
        & $setBrush 'chkIgnoreAllSpaces' 'Foreground' 'TextPrimary'

        # Inputs and pickers
        & $setBrush 'txtBasePath' 'Background' 'BgInput'
        & $setBrush 'txtBasePath' 'Foreground' 'TextPrimary'
        & $setBrush 'txtBasePath' 'BorderBrush' 'BorderInput'

        & $setBrush 'txtIncomingPath' 'Background' 'BgInput'
        & $setBrush 'txtIncomingPath' 'Foreground' 'TextPrimary'
        & $setBrush 'txtIncomingPath' 'BorderBrush' 'BorderInput'

        & $setBrush 'txtSearchReview' 'Background' 'BgInput'
        & $setBrush 'txtSearchReview' 'Foreground' 'TextPrimary'
        & $setBrush 'txtSearchReview' 'BorderBrush' 'BorderInput'

        & $setBrush 'txtSearchMapping' 'Foreground' 'TextPrimary'

        & $setBrush 'cmbBaseSheet' 'Background' 'BgInput'
        & $setBrush 'cmbBaseSheet' 'Foreground' 'TextPrimary'

        & $setBrush 'cmbIncomingSheet' 'Background' 'BgInput'
        & $setBrush 'cmbIncomingSheet' 'Foreground' 'TextPrimary'

        & $setBrush 'cmbLanguage' 'Background' 'BgInput'
        & $setBrush 'cmbLanguage' 'Foreground' 'TextPrimary'

        & $setBrush 'cmbFilterStatus' 'Background' 'BgInput'
        & $setBrush 'cmbFilterStatus' 'Foreground' 'TextPrimary'

        & $setBrush 'lbJoinBase' 'Background' 'BgInput'
        & $setBrush 'lbJoinBase' 'Foreground' 'TextPrimary'
        & $setBrush 'lbJoinBase' 'BorderBrush' 'BorderInput'

        & $setBrush 'lbJoinIncoming' 'Background' 'BgInput'
        & $setBrush 'lbJoinIncoming' 'Foreground' 'TextPrimary'
        & $setBrush 'lbJoinIncoming' 'BorderBrush' 'BorderInput'

        $gridMapRules = $w.FindName('gridMappingRules')
        if ($gridMapRules) {
            $gridMapRules.Background = $w.Resources['BgInput']
            $gridMapRules.Foreground = $w.Resources['TextPrimary']
            $gridMapRules.BorderBrush = $w.Resources['BorderCard']
            $gridMapRules.RowBackground = $w.Resources['DataGridRowBg']
            $gridMapRules.AlternatingRowBackground = $w.Resources['DataGridAltRowBg']
            $gridMapRules.HorizontalGridLinesBrush = $w.Resources['GridLines']
        }

        foreach ($dgName in @('dgBaseFullRow', 'gridMappingResultPreview', 'gridPrevIncomingRow', 'gridPrevBaseRow')) {
            $dg = $w.FindName($dgName)
            if ($dg) {
                $dg.Background = $w.Resources['BgInput']
                $dg.Foreground = $w.Resources['TextPrimary']
                $dg.BorderBrush = $w.Resources['BorderCard']
                $dg.RowBackground = $w.Resources['DataGridRowBg']
                $dg.AlternatingRowBackground = $w.Resources['DataGridAltRowBg']
                $dg.HorizontalGridLinesBrush = $w.Resources['GridLines']
                $dg.VerticalGridLinesBrush = $w.Resources['GridLines']
            }
        }

        & $setBrush 'lbReviewItems' 'Background' 'BgCard'
        & $setBrush 'lbReviewItems' 'BorderBrush' 'BorderCard'

        # Secondary action buttons
        $secBtnNames = @(
            'btnBrowseBase', 'btnBrowseIncoming', 'btnThemeToggle', 'btnSettings',
            'btnRestoreBackup', 'btnOpenBackups', 'btnOpenLogs', 'btnHelp',
            'btnAddRule', 'btnEditRule', 'btnRemoveRule', 'btnSaveProfile',
            'btnRefreshPreview', 'btnPrevSampleRow', 'btnNextSampleRow',
            'btnBackRow', 'btnSkipRow', 'btnExportReport', 'btnUndo'
        )
        foreach ($bn in $secBtnNames) {
            $b = $w.FindName($bn)
            if ($b) {
                $b.Background = $w.Resources['BtnSecondaryBg']
                $b.Foreground = $w.Resources['BtnSecondaryFg']
                $b.BorderBrush = $w.Resources['BorderCard']
                $b.BorderThickness = [System.Windows.Thickness]::new(1)
            }
        }

        # Update button text & DWM title bar chrome
        $bTheme = $w.FindName('btnThemeToggle')
        if ($bTheme) {
            $bTheme.Content = if ($isDark) { Get-UiString 'BtnThemeDark' } else { Get-UiString 'BtnThemeLight' }
        }
        try {
            $hw = (New-Object System.Windows.Interop.WindowInteropHelper($w)).Handle
            if ($hw -and $hw -ne [IntPtr]::Zero) {
                Set-WindowDwmTheme -Hwnd $hw -IsDark $isDark
            }
        } catch {}

        # Refresh review item badges to active theme
        $lbRev = $w.FindName('lbReviewItems')
        if ($script:AllReviewItems -and $lbRev) {
            foreach ($item in $script:AllReviewItems) {
                $stat = if ($item.Decision -and $item.Decision -ne 'Pending') { $item.Decision } else { $item.Record.Status }
                $bColors = Get-StatusBadgeColors -Status $stat -ThemeName $ThemeName
                $item.StatusBg = $bColors.Bg
                $item.StatusFg = $bColors.Fg
            }
            $lbRev.Items.Refresh()
        }

        # Refresh preview card badges to active theme
        if ($script:UpdateDataMappingPreview) {
            & $script:UpdateDataMappingPreview
        }

        # Re-render detail pane if an item is selected
        if ($RenderDetailPane -and $lbRev -and $lbRev.SelectedItem) {
            & $RenderDetailPane $lbRev.SelectedItem
        }

        if ($UpdateKpiPillSelection) {
            & $UpdateKpiPillSelection
        }
    }
    $script:ApplyTheme = $ApplyTheme

    # Helper: Update Staging Summary
    $UpdateStagingSummary = {
        $lblStaging = if ($txtStagingSummary) { $txtStagingSummary } elseif ($script:txtStagingSummary) { $script:txtStagingSummary } else { $null }
        if (-not $lblStaging) { return }
        if (-not $script:AllReviewItems -or $script:AllReviewItems.Count -eq 0) {
            $lblStaging.Text = ''
            return
        }
        $acceptedItems = @($script:AllReviewItems | Where-Object { $_.Decision -eq 'Accepted' })
        $stagedChanges = 0
        $skippedChanges = 0
        $affectedRows = 0

        foreach ($item in $acceptedItems) {
            $rec = $item.Record
            $hasStagedInRow = $false
            if ($rec.Changes -and $rec.Changes.Count -gt 0) {
                foreach ($chg in $rec.Changes) {
                    $isSelected = if ($item.SelectedCells) { ($item.SelectedCells[$chg.BaseColumn] -ne $false) } else { $true }
                    if ($isSelected) {
                        $stagedChanges++
                        $hasStagedInRow = $true
                    } else {
                        $skippedChanges++
                    }
                }
            } elseif ($rec.Status -eq 'New') {
                $stagedChanges++
                $hasStagedInRow = $true
            }
            if ($hasStagedInRow) { $affectedRows++ }
        }

        $fmt = Get-UiString 'StagingSummaryFormat' 'Wybrano do zapisu: {0} zmian w {1} wierszach (pominięto: {2})'
        $lblStaging.Text = $fmt -f $stagedChanges, $affectedRows, $skippedChanges

        $hasAccepted = $acceptedItems.Count -gt 0
        if ($btnApplyAccepted) { $btnApplyAccepted.IsEnabled = $hasAccepted }
        if ($btnApplyToNewFile) { $btnApplyToNewFile.IsEnabled = $hasAccepted }
    }
    $script:UpdateStagingSummary = $UpdateStagingSummary

    # Helper: Update Counters
    $UpdateCounters = {
        if (-not $script:AllReviewItems -or $script:AllReviewItems.Count -eq 0) {
            if ($script:txtCounters)   { $script:txtCounters.Text = (Get-UiString 'CountersFormat') -f 0, 0, 0, 0, 0, 0, 0 }
            elseif ($txtCounters)      { $txtCounters.Text = (Get-UiString 'CountersFormat') -f 0, 0, 0, 0, 0, 0, 0 }
            if ($txtKpiCountAll)       { $txtKpiCountAll.Text = "0" }
            if ($txtKpiCountNew)       { $txtKpiCountNew.Text = "0" }
            if ($txtKpiCountChanged)   { $txtKpiCountChanged.Text = "0" }
            if ($txtKpiCountAmbiguous) { $txtKpiCountAmbiguous.Text = "0" }
            if ($txtKpiCountAccepted)  { $txtKpiCountAccepted.Text = "0" }
            if ($txtKpiCountSkipped)   { $txtKpiCountSkipped.Text = "0" }
            if ($btnUndo)              { $btnUndo.IsEnabled = ($script:UndoStack.Count -gt 0) }
            & $UpdateStagingSummary
            return
        }
        $total = $script:AllReviewItems.Count
        $cNew  = ($script:AllReviewItems | Where-Object { $_.Record.Status -eq 'New' }).Count
        $cChg  = ($script:AllReviewItems | Where-Object { $_.Record.Status -eq 'Changed' }).Count
        $cAmb  = ($script:AllReviewItems | Where-Object { $_.Record.Status -eq 'Ambiguous' }).Count
        $cSkip = ($script:AllReviewItems | Where-Object { $_.Decision -eq 'Skipped' }).Count
        $cAcc  = ($script:AllReviewItems | Where-Object { $_.Decision -eq 'Accepted' }).Count
        $cRej  = ($script:AllReviewItems | Where-Object { $_.Decision -eq 'Rejected' }).Count
        if ($script:txtCounters)   { $script:txtCounters.Text = (Get-UiString 'CountersFormat') -f $total, $cNew, $cChg, $cAmb, $cSkip, $cAcc, $cRej }
        elseif ($txtCounters)      { $txtCounters.Text = (Get-UiString 'CountersFormat') -f $total, $cNew, $cChg, $cAmb, $cSkip, $cAcc, $cRej }
        if ($txtKpiCountAll)       { $txtKpiCountAll.Text = "$total" }
        if ($txtKpiCountNew)       { $txtKpiCountNew.Text = "$cNew" }
        if ($txtKpiCountChanged)   { $txtKpiCountChanged.Text = "$cChg" }
        if ($txtKpiCountAmbiguous) { $txtKpiCountAmbiguous.Text = "$cAmb" }
        if ($txtKpiCountAccepted)  { $txtKpiCountAccepted.Text = "$cAcc" }
        if ($txtKpiCountSkipped)   { $txtKpiCountSkipped.Text = "$($cSkip + $cRej)" }
        if ($btnUndo)              { $btnUndo.IsEnabled = ($script:UndoStack.Count -gt 0) }
        & $UpdateStagingSummary
    }
    $script:UpdateCounters = $UpdateCounters

    # Helper: Dynamic Localization
    $UpdateLocalization = {
        $w = if ($window) { $window } else { $script:window }
        if (-not $w) { return }
        $w.Title = Get-UiString 'AppTitle'

        $setProp = {
            param($name, $prop, $val)
            $ctrl = $w.FindName($name)
            if ($ctrl -and $null -ne $val) { $ctrl.$prop = $val }
        }

        & $setProp 'txtAppTitle' 'Text' (Get-UiString 'AppTitle')
        & $setProp 'txtAppSubtitle' 'Text' (Get-UiString 'AppSubtitle')
        & $setProp 'lblBasePath' 'Text' (Get-UiString 'LblBasePath')
        & $setProp 'lblBaseSheet' 'Text' (Get-UiString 'LblSheet')
        & $setProp 'btnBrowseBase' 'Content' (Get-UiString 'BtnBrowseBase')
        & $setProp 'lblIncomingPath' 'Text' (Get-UiString 'LblIncomingPath')
        & $setProp 'lblIncomingSheet' 'Text' (Get-UiString 'LblSheet')
        & $setProp 'btnBrowseIncoming' 'Content' (Get-UiString 'BtnBrowseIncoming')
        & $setProp 'lblLanguage' 'Text' (Get-UiString 'LblLanguage')
        & $setProp 'btnSettings' 'Content' (Get-UiString 'BtnSettings')
        & $setProp 'btnRestoreBackup' 'Content' (Get-UiString 'BtnRestoreBackup')
        & $setProp 'btnOpenBackups' 'Content' (Get-UiString 'BtnOpenBackups')
        & $setProp 'btnOpenBackups' 'ToolTip' (Get-UiString 'TooltipOpenBackups')
        & $setProp 'btnOpenLogs' 'Content' (Get-UiString 'BtnOpenLogs')
        & $setProp 'btnOpenLogs' 'ToolTip' (Get-UiString 'TooltipOpenLogs')
        & $setProp 'btnHelp' 'Content' (Get-UiString 'BtnHelp')
        & $setProp 'btnHelp' 'ToolTip' (Get-UiString 'TooltipHelp')
        & $setProp 'tabMapping' 'Header' (Get-UiString 'TabMapping')
        & $setProp 'btnAutoMap' 'Content' (Get-UiString 'BtnAutoMap')
        & $setProp 'btnAddRule' 'Content' (Get-UiString 'BtnAddRule')
        & $setProp 'btnEditRule' 'Content' (Get-UiString 'BtnEditRule')
        & $setProp 'btnRemoveRule' 'Content' (Get-UiString 'BtnRemoveRule')
        & $setProp 'btnSaveProfile' 'Content' (Get-UiString 'BtnSaveProfile')
        & $setProp 'colBase' 'Header' (Get-UiString 'ColBase')
        & $setProp 'colIncoming' 'Header' (Get-UiString 'ColIncoming')
        & $setProp 'colMergeMode' 'Header' (Get-UiString 'ColMergeMode')
        & $setProp 'colSeparator' 'Header' (Get-UiString 'ColSeparator')
        & $setProp 'colDiffPolicy' 'Header' (Get-UiString 'ColDiffPolicy')
        & $setProp 'lblJoinKeyTitle' 'Text' (Get-UiString 'LblJoinKeyTitle')
        & $setProp 'lblJoinKeyHint' 'Text' (Get-UiString 'LblJoinKeyHint')
        & $setProp 'lblJoinBaseSub' 'Text' (Get-UiString 'LblJoinBase')
        & $setProp 'lblJoinIncomingSub' 'Text' (Get-UiString 'LblJoinIncoming')
        & $setProp 'lblCompareOptionsTitle' 'Text' (Get-UiString 'LblCompareOptionsTitle')
        & $setProp 'chkIgnoreCase' 'Content' (Get-UiString 'LblIgnoreCase')
        & $setProp 'chkIgnoreCase' 'ToolTip' (Get-UiString 'TooltipIgnoreCase')
        & $setProp 'chkTrimWhitespace' 'Content' (Get-UiString 'LblTrimWhitespace')
        & $setProp 'chkTrimWhitespace' 'ToolTip' (Get-UiString 'TooltipTrimWhitespace')
        & $setProp 'chkIgnoreSpecialChars' 'Content' (Get-UiString 'LblIgnoreSpecialChars')
        & $setProp 'chkIgnoreSpecialChars' 'ToolTip' (Get-UiString 'TooltipIgnoreSpecialChars')
        & $setProp 'chkIgnoreAllSpaces' 'Content' (Get-UiString 'LblIgnoreAllSpaces')
        & $setProp 'chkIgnoreAllSpaces' 'ToolTip' (Get-UiString 'TooltipIgnoreAllSpaces')
        & $setProp 'btnRunCompare' 'Content' (Get-UiString 'BtnRunCompare')
        & $setProp 'tabReview' 'Header' (Get-UiString 'TabReview')
        & $setProp 'cbiFilterAll' 'Content' (Get-UiString 'FilterAll')
        & $setProp 'cbiFilterNew' 'Content' (Get-UiString 'FilterNew')
        & $setProp 'cbiFilterChanged' 'Content' (Get-UiString 'FilterChanged')
        & $setProp 'cbiFilterRemoved' 'Content' (Get-UiString 'FilterRemoved')
        & $setProp 'cbiFilterAmbiguous' 'Content' (Get-UiString 'FilterAmbiguous')
        & $setProp 'cbiFilterUnchanged' 'Content' (Get-UiString 'FilterUnchanged')
        & $setProp 'cbiFilterAccepted' 'Content' (Get-UiString 'FilterAccepted')
        & $setProp 'cbiFilterSkipped' 'Content' (Get-UiString 'FilterSkipped')
        & $setProp 'cbiFilterRejected' 'Content' (Get-UiString 'FilterRejected')
        & $setProp 'chkShowUnchanged' 'Content' (Get-UiString 'LblShowUnchanged')
        & $setProp 'chkShowUnchanged' 'ToolTip' (Get-UiString 'TooltipShowUnchanged')
        & $setProp 'btnBackRow' 'Content' (Get-UiString 'BtnBack')
        & $setProp 'btnAcceptRow' 'Content' (Get-UiString 'BtnAccept')
        & $setProp 'btnRejectRow' 'Content' (Get-UiString 'BtnReject')
        & $setProp 'btnSkipRow' 'Content' (Get-UiString 'BtnSkip')
        & $setProp 'btnEditRow' 'Content' (Get-UiString 'BtnEdit')
        & $setProp 'btnUndo' 'Content' (Get-UiString 'BtnUndo')
        & $setProp 'lblKpiAll' 'Text' (Get-UiString 'KpiAll')
        & $setProp 'lblKpiNew' 'Text' (Get-UiString 'KpiNew')
        & $setProp 'lblKpiChanged' 'Text' (Get-UiString 'KpiChanged')
        & $setProp 'lblKpiAmbiguous' 'Text' (Get-UiString 'KpiAmbiguous')
        & $setProp 'lblKpiAccepted' 'Text' (Get-UiString 'KpiAccepted')
        & $setProp 'lblKpiSkipped' 'Text' (Get-UiString 'KpiSkipped')
        & $setProp 'txtBasePath' 'ToolTip' (Get-UiString 'TooltipDropBase')
        & $setProp 'txtIncomingPath' 'ToolTip' (Get-UiString 'TooltipDropIncoming')
        & $setProp 'cardBaseFile' 'ToolTip' (Get-UiString 'TooltipDropBase')
        & $setProp 'cardIncomingFile' 'ToolTip' (Get-UiString 'TooltipDropIncoming')
        & $setProp 'txtSearchReview' 'ToolTip' (Get-UiString 'TooltipSearchReview')
        & $setProp 'btnSearchClear' 'ToolTip' (Get-UiString 'TooltipBtnSearchClear')
        & $setProp 'txtSearchMapping' 'ToolTip' (Get-UiString 'TooltipSearchMapping')
        & $setProp 'btnSearchMappingClear' 'ToolTip' (Get-UiString 'TooltipBtnSearchClear')
        & $setProp 'btnThemeToggle' 'ToolTip' (Get-UiString 'TooltipThemeToggle')
        & $setProp 'btnSettings' 'ToolTip' (Get-UiString 'TooltipSettings')
        & $setProp 'btnRestoreBackup' 'ToolTip' (Get-UiString 'TooltipRestoreBackup')
        & $setProp 'btnBrowseBase' 'ToolTip' (Get-UiString 'TooltipBrowseBase')
        & $setProp 'btnBrowseIncoming' 'ToolTip' (Get-UiString 'TooltipBrowseIncoming')
        & $setProp 'cmbBaseSheet' 'ToolTip' (Get-UiString 'TooltipCmbBaseSheet')
        & $setProp 'cmbIncomingSheet' 'ToolTip' (Get-UiString 'TooltipCmbIncomingSheet')
        & $setProp 'btnAutoMap' 'ToolTip' (Get-UiString 'TooltipAutoMap')
        & $setProp 'btnAddRule' 'ToolTip' (Get-UiString 'TooltipAddRule')
        & $setProp 'btnEditRule' 'ToolTip' (Get-UiString 'TooltipEditRule')
        & $setProp 'btnRemoveRule' 'ToolTip' (Get-UiString 'TooltipRemoveRule')
        & $setProp 'btnSaveProfile' 'ToolTip' (Get-UiString 'TooltipSaveProfile')
        & $setProp 'lbJoinBase' 'ToolTip' (Get-UiString 'TooltipJoinBase')
        & $setProp 'lbJoinIncoming' 'ToolTip' (Get-UiString 'TooltipJoinIncoming')
        & $setProp 'btnRunCompare' 'ToolTip' (Get-UiString 'TooltipRunCompare')
        & $setProp 'lblBatchColToggles' 'Text' (Get-UiString 'LblBatchColumnToggles')

        if ($script:MappingRules) {
            foreach ($r in $script:MappingRules) {
                if ($r.PSObject.Properties['DiffPolicy']) {
                    $r.DiffPolicyDisplay = Get-DiffPolicyDisplay $r.DiffPolicy
                }
            }
            if ($gridMappingRules) { $gridMappingRules.Items.Refresh() }
        }
        & $UpdateStagingSummary
        & $setProp 'cmbFilterStatus' 'ToolTip' (Get-UiString 'TooltipFilterStatus')
        & $setProp 'btnBackRow' 'ToolTip' (Get-UiString 'TooltipBtnBack')
        & $setProp 'btnAcceptRow' 'ToolTip' (Get-UiString 'TooltipBtnAccept')
        & $setProp 'btnRejectRow' 'ToolTip' (Get-UiString 'TooltipBtnReject')
        & $setProp 'btnSkipRow' 'ToolTip' (Get-UiString 'TooltipBtnSkip')
        & $setProp 'btnEditRow' 'ToolTip' (Get-UiString 'TooltipBtnEdit')
        & $setProp 'btnUndo' 'ToolTip' (Get-UiString 'TooltipBtnUndo')
        & $setProp 'btnAcceptAll' 'ToolTip' (Get-UiString 'TooltipBtnAcceptAll')
        & $setProp 'btnRejectAll' 'ToolTip' (Get-UiString 'TooltipBtnRejectAll')
        & $setProp 'btnExportReport' 'ToolTip' (Get-UiString 'TooltipBtnExportReport')
        & $setProp 'btnApplyAccepted' 'ToolTip' (Get-UiString 'TooltipBtnApplyAccepted')
        & $setProp 'btnApplyToNewFile' 'ToolTip' (Get-UiString 'TooltipBtnApplyToNewFile')
        & $setProp 'btnModifyBaseField' 'ToolTip' (Get-UiString 'TooltipBtnModifyBaseField')
        & $setProp 'cmbLanguage' 'ToolTip' (Get-UiString 'TooltipCmbLanguage')
        & $setProp 'btnAcceptAll' 'Content' (Get-UiString 'BtnAcceptAll')
        & $setProp 'btnRejectAll' 'Content' (Get-UiString 'BtnRejectAll')
        & $setProp 'btnExportReport' 'Content' (Get-UiString 'BtnExportReport')
        & $setProp 'btnApplyAccepted' 'Content' (Get-UiString 'BtnApplyAccepted')
        & $setProp 'btnApplyToNewFile' 'Content' (Get-UiString 'BtnApplyToNewFile')
        & $setProp 'txtBaseRowHint' 'Text' (Get-UiString 'BaseRowHint')
        & $setProp 'btnModifyBaseField' 'Content' (Get-UiString 'BtnModifyBaseField')
        & $setProp 'txtPreviewHeader' 'Text' (Get-UiString 'PreviewCardTitle')
        & $setProp 'txtPreviewHint' 'Text' (Get-UiString 'PreviewCardHint')
        & $setProp 'tabPrevResult' 'Header' (Get-UiString 'TabPrevResult')
        & $setProp 'tabPrevIncoming' 'Header' (Get-UiString 'TabPrevIncoming')
        & $setProp 'tabPrevBase' 'Header' (Get-UiString 'TabPrevBase')
        & $setProp 'btnRefreshPreview' 'Content' (Get-UiString 'BtnRefreshPreview')
        & $setProp 'btnPrevSampleRow' 'Content' (Get-UiString 'BtnPrevSampleRow')
        & $setProp 'btnNextSampleRow' 'Content' (Get-UiString 'BtnNextSampleRow')
        & $setProp 'colPrevTargetBase' 'Header' (Get-UiString 'ColPrevTargetBase')
        & $setProp 'colPrevSourceExpr' 'Header' (Get-UiString 'ColPrevSourceExpr')
        & $setProp 'colPrevProjectedVal' 'Header' (Get-UiString 'ColPrevProjectedVal')
        & $setProp 'colPrevCurrentBaseVal' 'Header' (Get-UiString 'ColPrevCurrentBaseVal')
        & $setProp 'colPrevStatus' 'Header' (Get-UiString 'ColPrevStatus')
        & $setProp 'txtEmptyPreviewPrompt' 'Text' (Get-UiString 'PreviewLoadFilesPrompt')

        if ($ApplyTheme) { & $ApplyTheme $script:CurrentTheme }
        elseif ($script:ApplyTheme) { & $script:ApplyTheme $script:CurrentTheme }
        if ($UpdateCounters) { & $UpdateCounters }
        elseif ($script:UpdateCounters) { & $script:UpdateCounters }

        $tDetail = $w.FindName('txtDetailHeader')
        $lbRev = $w.FindName('lbReviewItems')
        if ($tDetail -and $lbRev -and -not $lbRev.SelectedItem) {
            $tDetail.Text = Get-UiString 'SelectRowHeader'
        }
    }
    $script:UpdateLocalization = $UpdateLocalization

    # Populate Language ComboBox
    foreach ($code in @('pl', 'en', 'de')) {
        $langObj = $script:LanguagesCatalog[$code]
        $cbi = [System.Windows.Controls.ComboBoxItem]::new()
        $cbi.Content = if ($langObj) { $langObj.DisplayName } else { $code.ToUpper() }
        $cbi.Tag = $code
        [void]$cmbLanguage.Items.Add($cbi)
        if ($code -eq $script:CurrentLanguage) {
            $cmbLanguage.SelectedItem = $cbi
        }
    }
    if (-not $cmbLanguage.SelectedItem -and $cmbLanguage.Items.Count -gt 0) {
        $cmbLanguage.SelectedIndex = 0
    }

    $cmbLanguage.add_SelectionChanged({
        $selTag = if ($cmbLanguage.SelectedItem -is [System.Windows.Controls.ComboBoxItem]) {
            $cmbLanguage.SelectedItem.Tag
        } else { 'pl' }
        $script:CurrentLanguage = $selTag
        $script:AppConfig.UiLanguage = $selTag
        Save-AppConfig -Config $script:AppConfig
        & $script:UpdateLocalization
    })

    # Theme Toggle Handler
    $btnThemeToggle.add_Click({
        $newTheme = if ($script:CurrentTheme -eq 'Dark') { 'Light' } else { 'Dark' }
        $script:AppConfig.Theme = $newTheme
        Save-AppConfig -Config $script:AppConfig
        & $script:ApplyTheme $newTheme
    })

    # Check Profile Auto-Match
    $CheckProfileAutoMatch = {
        if (-not $script:IncomingHeaders -or $script:IncomingHeaders.Length -eq 0) { return }
        $sheetName = if ($cmbIncomingSheet.SelectedItem) { $cmbIncomingSheet.SelectedItem.ToString() } else { '' }
        $fp = Compute-HeaderFingerprint -Headers $script:IncomingHeaders -SheetName $sheetName
        $fpFallback = Compute-HeaderFingerprint -Headers $script:IncomingHeaders

        $prof = Get-MappingProfileByFingerprint -Fingerprint $fp -StorePath $script:AppConfig.ProfileStorePath
        if (-not $prof) {
            $prof = Get-MappingProfileByFingerprint -Fingerprint $fpFallback -StorePath $script:AppConfig.ProfileStorePath
        }

        if ($prof) {
            $script:CurrentProfile = $prof
            $badgeProfile.Visibility = [System.Windows.Visibility]::Visible
            $txtProfileBadge.Text = "$((Get-UiString 'BadgeProfileAuto')): $($prof.Name)"

            $script:MappingRules.Clear()
            foreach ($r in $prof.MappingRules) {
                $bCols = @($r.BaseColumns)
                $uCols = @($r.UpdateColumns)
                $dp = if ($r.DiffPolicy) { $r.DiffPolicy } else { 'TrackChanges' }
                $script:MappingRules.Add([PSCustomObject]@{
                    BaseColumns       = $bCols
                    UpdateColumns     = $uCols
                    MergeMode         = if ($r.MergeMode) { $r.MergeMode } else { 'Exact' }
                    Separator         = if ($r.Separator) { $r.Separator } else { '' }
                    DiffPolicy        = $dp
                    DiffPolicyDisplay = Get-DiffPolicyDisplay $dp
                    BaseColsStr       = $bCols -join ', '
                    UpdColsStr        = $uCols -join ', '
                })
            }

            # Select join keys
            if ($prof.JoinKeyBase) {
                $lbJoinBase.SelectedItems.Clear()
                foreach ($jk in $prof.JoinKeyBase) {
                    if ($lbJoinBase.Items.Contains($jk)) { [void]$lbJoinBase.SelectedItems.Add($jk) }
                }
            }
            if ($prof.JoinKeyUpdate) {
                $lbJoinIncoming.SelectedItems.Clear()
                foreach ($jk in $prof.JoinKeyUpdate) {
                    if ($lbJoinIncoming.Items.Contains($jk)) { [void]$lbJoinIncoming.SelectedItems.Add($jk) }
                }
            }
            if ($prof.CompareOptions) {
                $co = $prof.CompareOptions
                $coCase = if ($co -is [System.Collections.IDictionary]) { if ($co.Contains('IgnoreCase')) { [bool]$co['IgnoreCase'] } else { $null } } elseif ($co.PSObject.Properties['IgnoreCase']) { [bool]$co.IgnoreCase } else { $null }
                $coTrim = if ($co -is [System.Collections.IDictionary]) { if ($co.Contains('Trim')) { [bool]$co['Trim'] } elseif ($co.Contains('TrimWhitespace')) { [bool]$co['TrimWhitespace'] } else { $null } } elseif ($co.PSObject.Properties['Trim']) { [bool]$co.Trim } elseif ($co.PSObject.Properties['TrimWhitespace']) { [bool]$co.TrimWhitespace } else { $null }
                $coSpec = if ($co -is [System.Collections.IDictionary]) { if ($co.Contains('IgnoreSpecialChars')) { [bool]$co['IgnoreSpecialChars'] } else { $null } } elseif ($co.PSObject.Properties['IgnoreSpecialChars']) { [bool]$co.IgnoreSpecialChars } else { $null }
                $coSpc  = if ($co -is [System.Collections.IDictionary]) { if ($co.Contains('IgnoreAllSpaces')) { [bool]$co['IgnoreAllSpaces'] } else { $null } } elseif ($co.PSObject.Properties['IgnoreAllSpaces']) { [bool]$co.IgnoreAllSpaces } else { $null }

                if ($null -ne $coCase -and $script:chkIgnoreCase) { $script:chkIgnoreCase.IsChecked = $coCase }
                if ($null -ne $coTrim -and $script:chkTrimWhitespace) { $script:chkTrimWhitespace.IsChecked = $coTrim }
                if ($null -ne $coSpec -and $script:chkIgnoreSpecialChars) { $script:chkIgnoreSpecialChars.IsChecked = $coSpec }
                if ($null -ne $coSpc -and $script:chkIgnoreAllSpaces) { $script:chkIgnoreAllSpaces.IsChecked = $coSpc }
            }
            $txtStatusMsg.Text = (Get-UiString 'ProfileSavedMsg') -f $prof.Name
        } else {
            $badgeProfile.Visibility = [System.Windows.Visibility]::Collapsed
            # Fallback when no saved profile matches: auto-map columns and select shared join key
            if ($script:BaseHeaders -and $script:IncomingHeaders -and $script:MappingRules.Count -eq 0) {
                $autoRules = Invoke-AutoMapRules -BaseHeaders $script:BaseHeaders -IncomingHeaders $script:IncomingHeaders
                foreach ($r in $autoRules) {
                    $script:MappingRules.Add($r)
                }
                $lbJoinBase.SelectedItems.Clear()
                $lbJoinIncoming.SelectedItems.Clear()
                $foundJoinKey = $false
                foreach ($ih in $script:IncomingHeaders) {
                    if ($ih -match '(?i)(id|kod|pesel|nr)') {
                        $normIh = $ih.Trim().ToLowerInvariant() -replace '[_\-\s]+', ''
                        $matchBase = $script:BaseHeaders | Where-Object {
                            $normBh = $_.Trim().ToLowerInvariant() -replace '[_\-\s]+', ''
                            $normBh -eq $normIh
                        } | Select-Object -First 1
                        if ($matchBase -and $lbJoinBase.Items.Contains($matchBase) -and $lbJoinIncoming.Items.Contains($ih)) {
                            [void]$lbJoinBase.SelectedItems.Add($matchBase)
                            [void]$lbJoinIncoming.SelectedItems.Add($ih)
                            $foundJoinKey = $true
                            break
                        }
                    }
                }
                if (-not $foundJoinKey) {
                    foreach ($ih in $script:IncomingHeaders) {
                        if ($ih -match '(?i)(nazwisk|imie|imię|dzieck|klient|uczen|uczeń|pracownik|nazwa|name|surname|firstname|lastname|child|student|employee|client|customer|nachname|vorname|kind|schueler|mitarbeiter|kunde)') {
                            $normIh = $ih.Trim().ToLowerInvariant() -replace '[_\-\s]+', ''
                            $matchBase = $script:BaseHeaders | Where-Object {
                                $normBh = $_.Trim().ToLowerInvariant() -replace '[_\-\s]+', ''
                                $normBh -eq $normIh
                            } | Select-Object -First 1
                            if ($matchBase -and $lbJoinBase.Items.Contains($matchBase) -and $lbJoinIncoming.Items.Contains($ih)) {
                                [void]$lbJoinBase.SelectedItems.Add($matchBase)
                                [void]$lbJoinIncoming.SelectedItems.Add($ih)
                                $foundJoinKey = $true
                                break
                            }
                        }
                    }
                }
                if (-not $foundJoinKey -and $script:IncomingHeaders -and $script:IncomingHeaders.Count -gt 0) {
                    $firstIh = $script:IncomingHeaders[0]
                    $normFirstIh = $firstIh.Trim().ToLowerInvariant() -replace '[_\-\s]+', ''
                    $matchBase = $script:BaseHeaders | Where-Object {
                        $normBh = $_.Trim().ToLowerInvariant() -replace '[_\-\s]+', ''
                        $normBh -eq $normFirstIh
                    } | Select-Object -First 1
                    if ($matchBase -and $lbJoinBase.Items.Contains($matchBase) -and $lbJoinIncoming.Items.Contains($firstIh)) {
                        [void]$lbJoinBase.SelectedItems.Add($matchBase)
                        [void]$lbJoinIncoming.SelectedItems.Add($firstIh)
                    }
                }
            }
        }
    }

    # Data & Mapping Live Preview Logic
    $script:PreviewRowIndex = 0
    $script:CachedBaseSampleRows = $null
    $script:CachedIncomingSampleRows = $null
    $script:CachedBaseFilePath = ''
    $script:CachedBaseSheet = ''
    $script:CachedIncomingFilePath = ''
    $script:CachedIncomingSheet = ''

    $script:BindSampleRowToGrid = {
        param($DataGrid, [System.Collections.IList]$Headers, $RowObj)
        if (-not $DataGrid -or -not $Headers -or $Headers.Count -eq 0 -or -not $RowObj) { return }

        $dt = New-Object System.Data.DataTable
        $DataGrid.Columns.Clear()

        # Add physical row column
        [void]$dt.Columns.Add('_RowNumber', [int])
        $colR = New-Object System.Windows.Controls.DataGridTextColumn -Property @{
            Header     = '#'
            Binding    = New-Object System.Windows.Data.Binding('[_RowNumber]')
            IsReadOnly = $true
            Width      = 45
        }
        $DataGrid.Columns.Add($colR)

        for ($i = 0; $i -lt $Headers.Count; $i++) {
            $h = $Headers[$i]
            $colKey = "Col_$i"
            [void]$dt.Columns.Add($colKey, [string])
            $col = New-Object System.Windows.Controls.DataGridTextColumn -Property @{
                Header     = $h
                Binding    = New-Object System.Windows.Data.Binding("[$colKey]")
                IsReadOnly = $true
            }
            $DataGrid.Columns.Add($col)
        }

        $dr = $dt.NewRow()
        $dr['_RowNumber'] = if ($RowObj.PSObject.Properties['_RowNumber'] -and $null -ne $RowObj.PSObject.Properties['_RowNumber'].Value) {
            [int]$RowObj.PSObject.Properties['_RowNumber'].Value
        } else {
            $script:PreviewRowIndex + 2
        }
        for ($i = 0; $i -lt $Headers.Count; $i++) {
            $h = $Headers[$i]
            $val = ''
            $prop = $RowObj.PSObject.Properties[$h]
            if ($prop -and $null -ne $prop.Value) {
                $val = $prop.Value.ToString()
            }
            $dr["Col_$i"] = $val
        }
        $dt.Rows.Add($dr)
        $DataGrid.ItemsSource = $dt.DefaultView
    }
    $BindSampleRowToGrid = $script:BindSampleRowToGrid

    $script:UpdateDataMappingPreview = {
        param([bool]$ForceReload = $false)
        if (-not $script:window) { return }

        if ($ForceReload) {
            $script:PreviewRowIndex = 0
        }

        $basePath = if ($script:txtBasePath) { $script:txtBasePath.Text } else { '' }
        $baseSheet = if ($script:cmbBaseSheet -and $script:cmbBaseSheet.SelectedItem) { $script:cmbBaseSheet.SelectedItem.ToString() } else { '' }
        $incPath = if ($script:txtIncomingPath) { $script:txtIncomingPath.Text } else { '' }
        $incSheet = if ($script:cmbIncomingSheet -and $script:cmbIncomingSheet.SelectedItem) { $script:cmbIncomingSheet.SelectedItem.ToString() } else { '' }

        $hasBase = (-not [string]::IsNullOrWhiteSpace($basePath)) -and (Test-Path $basePath) -and (-not [string]::IsNullOrEmpty($baseSheet))
        $hasInc = (-not [string]::IsNullOrWhiteSpace($incPath)) -and (Test-Path $incPath) -and (-not [string]::IsNullOrEmpty($incSheet))

        if ($hasBase) {
            try {
                if ($ForceReload -or -not $script:CachedBaseSampleRows -or $script:CachedBaseFilePath -ne $basePath -or $script:CachedBaseSheet -ne $baseSheet) {
                    $script:CachedBaseSampleRows = [FastExcelHelper]::ReadSheet($basePath, $baseSheet, 10)
                    $script:CachedBaseFilePath = $basePath
                    $script:CachedBaseSheet = $baseSheet
                }
                if ($script:CachedBaseSampleRows -and $script:CachedBaseSampleRows.Count -gt 0) {
                    $bIdx = [Math]::Min($script:PreviewRowIndex, $script:CachedBaseSampleRows.Count - 1)
                    $bRow = $script:CachedBaseSampleRows[$bIdx]
                    & $script:BindSampleRowToGrid $script:gridPrevBaseRow $script:BaseHeaders $bRow
                    if ($script:txtEmptyBasePrompt) { $script:txtEmptyBasePrompt.Visibility = [System.Windows.Visibility]::Collapsed }
                }
            } catch { }
        } else {
            if ($script:gridPrevBaseRow) { $script:gridPrevBaseRow.ItemsSource = $null }
            if ($script:txtEmptyBasePrompt) { $script:txtEmptyBasePrompt.Visibility = [System.Windows.Visibility]::Visible }
        }

        if ($hasInc) {
            try {
                if ($ForceReload -or -not $script:CachedIncomingSampleRows -or $script:CachedIncomingFilePath -ne $incPath -or $script:CachedIncomingSheet -ne $incSheet) {
                    $script:CachedIncomingSampleRows = [FastExcelHelper]::ReadSheet($incPath, $incSheet, 10)
                    $script:CachedIncomingFilePath = $incPath
                    $script:CachedIncomingSheet = $incSheet
                }
                if ($script:CachedIncomingSampleRows -and $script:CachedIncomingSampleRows.Count -gt 0) {
                    $iIdx = [Math]::Min($script:PreviewRowIndex, $script:CachedIncomingSampleRows.Count - 1)
                    $iRow = $script:CachedIncomingSampleRows[$iIdx]
                    & $script:BindSampleRowToGrid $script:gridPrevIncomingRow $script:IncomingHeaders $iRow
                    if ($script:txtEmptyIncomingPrompt) { $script:txtEmptyIncomingPrompt.Visibility = [System.Windows.Visibility]::Collapsed }
                }
            } catch { }
        } else {
            if ($script:gridPrevIncomingRow) { $script:gridPrevIncomingRow.ItemsSource = $null }
            if ($script:txtEmptyIncomingPrompt) { $script:txtEmptyIncomingPrompt.Visibility = [System.Windows.Visibility]::Visible }
        }

        if (-not $hasBase -or -not $hasInc) {
            if ($script:txtEmptyPreviewPrompt) { $script:txtEmptyPreviewPrompt.Visibility = [System.Windows.Visibility]::Visible }
            if ($script:gridMappingResultPreview) { $script:gridMappingResultPreview.ItemsSource = $null }
            if ($script:txtPreviewInfo) { $script:txtPreviewInfo.Text = '' }
            return
        }

        # Both files loaded!
        if ($script:txtEmptyPreviewPrompt) { $script:txtEmptyPreviewPrompt.Visibility = [System.Windows.Visibility]::Collapsed }

        try {
            if (-not $script:CachedIncomingSampleRows -or $script:CachedIncomingSampleRows.Count -eq 0 -or
                -not $script:CachedBaseSampleRows -or $script:CachedBaseSampleRows.Count -eq 0) {
                return
            }

            $bIdx = [Math]::Min($script:PreviewRowIndex, $script:CachedBaseSampleRows.Count - 1)
            $iIdx = [Math]::Min($script:PreviewRowIndex, $script:CachedIncomingSampleRows.Count - 1)

            $bRow = $script:CachedBaseSampleRows[$bIdx]
            $iRow = $script:CachedIncomingSampleRows[$iIdx]

            # Project incoming row through mapping rules
            $incHash = @{}
            foreach ($prop in $iRow.PSObject.Properties) {
                if ($prop.Name -ne '_RowNumber') {
                    $incHash[$prop.Name] = if ($null -ne $prop.Value) { $prop.Value.ToString() } else { '' }
                }
            }

            $baseHash = @{}
            foreach ($prop in $bRow.PSObject.Properties) {
                if ($prop.Name -ne '_RowNumber') {
                    $baseHash[$prop.Name] = if ($null -ne $prop.Value) { $prop.Value.ToString() } else { '' }
                }
            }

            $optTrimActive = if ($script:chkTrimWhitespace) { [bool]$script:chkTrimWhitespace.IsChecked } else { $true }
            $projected = Get-ProjectedRow -IncomingValues $incHash -MappingRules $script:MappingRules -Trim $optTrimActive

            # Build Mapping Results Preview Items
            $previewList = [System.Collections.Generic.List[object]]::new()
            $mappedCount = 0

            $baseHeadersList = if ($script:BaseHeaders) { $script:BaseHeaders } else { @($baseHash.Keys) }
            $selectedBaseKeys = if ($script:lbJoinBase -and $script:lbJoinBase.SelectedItems) {
                @($script:lbJoinBase.SelectedItems | ForEach-Object { $_.ToString() })
            } else { @() }
            foreach ($bh in $baseHeadersList) {
                $isJoinKey = ($selectedBaseKeys -contains $bh)
                $rule = $null
                foreach ($r in $script:MappingRules) {
                    if ($r.BaseColumns -and $r.BaseColumns -contains $bh) {
                        $rule = $r
                        break
                    }
                }
                $isMapped = ($null -ne $rule)
                $sourceExpr = if ($isMapped) {
                    $uCols = if ($rule.UpdColsStr) { $rule.UpdColsStr } else { [string]::Join(', ', $rule.UpdateColumns) }
                    if ($rule.MergeMode -and $rule.MergeMode -ne 'Exact') {
                        "$uCols ($($rule.MergeMode))"
                    } else {
                        $uCols
                    }
                } else {
                    "—"
                }

                $projVal = if ($isMapped -and $projected.ContainsKey($bh)) { $projected[$bh] } else { '' }
                $baseVal = if ($baseHash.ContainsKey($bh)) { $baseHash[$bh] } else { '' }

                if ($isMapped) { $mappedCount++ }

                # Determine Preview Status & Colors (Theme-aware)
                $statusText = ''
                $isLight = ($script:CurrentTheme -eq 'Light')
                $statusBg = if ($isLight) { '#F1F5F9' } else { '#334155' }
                $statusFg = if ($isLight) { '#475569' } else { '#94A3B8' }

                $optCase = if ($script:chkIgnoreCase) { [bool]$script:chkIgnoreCase.IsChecked } else { $true }
                $optTrim = if ($script:chkTrimWhitespace) { [bool]$script:chkTrimWhitespace.IsChecked } else { $true }
                $optSpec = if ($script:chkIgnoreSpecialChars) { [bool]$script:chkIgnoreSpecialChars.IsChecked } else { $true }
                $optSpc  = if ($script:chkIgnoreAllSpaces) { [bool]$script:chkIgnoreAllSpaces.IsChecked } else { $true }

                if (-not $isMapped) {
                    $statusText = Get-UiString 'PreviewStatusUnmapped'
                    $statusBg = if ($isLight) { '#F1F5F9' } else { '#1E293B' }
                    $statusFg = if ($isLight) { '#475569' } else { '#64748B' }
                } elseif ([FastDiffHelper]::AreEqual($baseVal, $projVal, $optCase, $optTrim, $optSpec, $optSpc)) {
                    $statusText = Get-UiString 'PreviewStatusIdentical'
                    $statusBg = if ($isLight) { '#D1FAE5' } else { '#064E3B' }
                    $statusFg = if ($isLight) { '#065F46' } else { '#34D399' }
                } elseif ([string]::IsNullOrWhiteSpace($baseVal) -and -not [string]::IsNullOrWhiteSpace($projVal)) {
                    $statusText = Get-UiString 'PreviewStatusNewValue'
                    $statusBg = if ($isLight) { '#DBEAFE' } else { '#1E3A8A' }
                    $statusFg = if ($isLight) { '#1E40AF' } else { '#60A5FA' }
                } else {
                    $statusText = Get-UiString 'PreviewStatusChanged'
                    $statusBg = if ($isLight) { '#FEF3C7' } else { '#713F12' }
                    $statusFg = if ($isLight) { '#92400E' } else { '#FBBF24' }
                }

                $previewList.Add([PSCustomObject]@{
                    TargetBaseColumn    = if ($isJoinKey) { "🔑 $bh" } else { $bh }
                    MappedSourceColumns = $sourceExpr
                    ProjectedValue      = $projVal
                    CurrentBaseValue    = $baseVal
                    StatusText          = $statusText
                    StatusBg            = $statusBg
                    StatusFg            = $statusFg
                    IsJoinKey           = $isJoinKey
                })
            }

            $script:gridMappingResultPreview.ItemsSource = $previewList

            $incHeadersCount = if ($script:IncomingHeaders) { $script:IncomingHeaders.Count } else { $incHash.Count }
            if ($script:txtPreviewInfo) {
                $script:txtPreviewInfo.Text = (Get-UiString 'PreviewRowIndicator') -f ($iIdx + 1), $baseHeadersList.Count, $incHeadersCount, $mappedCount, $baseHeadersList.Count
            }

            $maxAvailable = 0
            if ($script:CachedIncomingSampleRows) { $maxAvailable = [Math]::Max($maxAvailable, $script:CachedIncomingSampleRows.Count) }
            if ($script:CachedBaseSampleRows) { $maxAvailable = [Math]::Max($maxAvailable, $script:CachedBaseSampleRows.Count) }

            if ($script:btnPrevSampleRow) {
                $script:btnPrevSampleRow.IsEnabled = ($script:PreviewRowIndex -gt 0)
            }
            if ($script:btnNextSampleRow) {
                $script:btnNextSampleRow.IsEnabled = ($maxAvailable -gt 0 -and $script:PreviewRowIndex -lt ($maxAvailable - 1))
            }
        } catch {
            if ($script:txtStatusMsg) { $script:txtStatusMsg.Text = (Get-UiString 'PreviewError') -f $_.Exception.Message }
        }
    }
    $UpdateDataMappingPreview = $script:UpdateDataMappingPreview

    # Comparison options toggle handlers: immediately refresh live mapping preview
    $onCompareOptionChanged = {
        if ($script:UpdateDataMappingPreview) {
            & $script:UpdateDataMappingPreview
        }
    }
    if ($chkIgnoreCase) {
        $chkIgnoreCase.add_Checked($onCompareOptionChanged)
        $chkIgnoreCase.add_Unchecked($onCompareOptionChanged)
    }
    if ($chkTrimWhitespace) {
        $chkTrimWhitespace.add_Checked($onCompareOptionChanged)
        $chkTrimWhitespace.add_Unchecked($onCompareOptionChanged)
    }
    if ($chkIgnoreSpecialChars) {
        $chkIgnoreSpecialChars.add_Checked($onCompareOptionChanged)
        $chkIgnoreSpecialChars.add_Unchecked($onCompareOptionChanged)
    }
    if ($chkIgnoreAllSpaces) {
        $chkIgnoreAllSpaces.add_Checked($onCompareOptionChanged)
        $chkIgnoreAllSpaces.add_Unchecked($onCompareOptionChanged)
    }

    # Load Base File Handler
    $LoadBaseFile = {
        param([string]$Path)
        if ([string]::IsNullOrWhiteSpace($Path)) { return }
        $cleanPath = $Path.Trim().Trim('"').Trim("'")
        if (-not (Test-Path $cleanPath)) {
            $msg = (Get-UiString 'ErrFileNotExist') -f $cleanPath
            $txtStatusMsg.Text = $msg
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'ErrorTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        if ((Get-Item $cleanPath).PSIsContainer) {
            $msg = (Get-UiString 'ErrPathIsFolder') -f $cleanPath
            $txtStatusMsg.Text = $msg
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $ext = [System.IO.Path]::GetExtension($cleanPath).ToLowerInvariant()
        if ($ext -notin @('.xlsx', '.csv')) {
            $msg = (Get-UiString 'ErrUnsupportedFormat') -f $ext
            $txtStatusMsg.Text = $msg
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        if ((Get-Item $cleanPath).Length -eq 0) {
            $msg = (Get-UiString 'ErrEmptyFile') -f [System.IO.Path]::GetFileName($cleanPath)
            $txtStatusMsg.Text = $msg
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $txtBasePath.Text = $cleanPath
        if ($script:AppConfig.RememberBasePath -ne $false) {
            $script:AppConfig.BaseFilePath = $cleanPath
            Save-AppConfig -Config $script:AppConfig
        }
        $cmbBaseSheet.Items.Clear()
        try {
            $sheets = [FastExcelHelper]::GetSheetNames($cleanPath)
            foreach ($s in $sheets) { [void]$cmbBaseSheet.Items.Add($s) }
            if ($cmbBaseSheet.Items.Count -gt 0) {
                $selIdx = 0
                if (-not [string]::IsNullOrEmpty($script:AppConfig.BaseSheet)) {
                    $idx = $sheets.IndexOf($script:AppConfig.BaseSheet)
                    if ($idx -ge 0) { $selIdx = $idx }
                }
                $cmbBaseSheet.SelectedIndex = $selIdx
            }
            $txtStatusMsg.Text = (Get-UiString 'StatusLoadedBase') -f [System.IO.Path]::GetFileName($cleanPath), $cmbBaseSheet.Items.Count
            & $script:UpdateDataMappingPreview $true
        } catch {
            $msg = (Get-UiString 'ErrFileLockedOrCorrupt') -f [System.IO.Path]::GetFileName($cleanPath), $_.Exception.Message
            $txtStatusMsg.Text = (Get-UiString 'ErrReadBase') -f $_.Exception.Message
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'ErrorTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }

    # Load Incoming File Handler
    $LoadIncomingFile = {
        param([string]$Path)
        if ([string]::IsNullOrWhiteSpace($Path)) { return }
        $cleanPath = $Path.Trim().Trim('"').Trim("'")
        if (-not (Test-Path $cleanPath)) {
            $msg = (Get-UiString 'ErrFileNotExist') -f $cleanPath
            $txtStatusMsg.Text = $msg
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'ErrorTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        if ((Get-Item $cleanPath).PSIsContainer) {
            $msg = (Get-UiString 'ErrPathIsFolder') -f $cleanPath
            $txtStatusMsg.Text = $msg
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $ext = [System.IO.Path]::GetExtension($cleanPath).ToLowerInvariant()
        if ($ext -notin @('.xlsx', '.csv')) {
            $msg = (Get-UiString 'ErrUnsupportedFormat') -f $ext
            $txtStatusMsg.Text = $msg
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        if ((Get-Item $cleanPath).Length -eq 0) {
            $msg = (Get-UiString 'ErrEmptyFile') -f [System.IO.Path]::GetFileName($cleanPath)
            $txtStatusMsg.Text = $msg
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $txtIncomingPath.Text = $cleanPath
        $cmbIncomingSheet.Items.Clear()
        try {
            $sheets = [FastExcelHelper]::GetSheetNames($cleanPath)
            foreach ($s in $sheets) { [void]$cmbIncomingSheet.Items.Add($s) }
            if ($cmbIncomingSheet.Items.Count -gt 0) {
                $selIdx = 0
                if (-not [string]::IsNullOrEmpty($InitIncomingSheet)) {
                    $idx = $sheets.IndexOf($InitIncomingSheet)
                    if ($idx -ge 0) { $selIdx = $idx }
                }
                $cmbIncomingSheet.SelectedIndex = $selIdx
            }
            $txtStatusMsg.Text = (Get-UiString 'StatusLoadedIncoming') -f [System.IO.Path]::GetFileName($cleanPath), $cmbIncomingSheet.Items.Count
            & $script:UpdateDataMappingPreview $true
        } catch {
            $msg = (Get-UiString 'ErrFileLockedOrCorrupt') -f [System.IO.Path]::GetFileName($cleanPath), $_.Exception.Message
            $txtStatusMsg.Text = (Get-UiString 'ErrReadIncoming') -f $_.Exception.Message
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'ErrorTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }

    # Sheet selection change handlers
    $cmbBaseSheet.add_SelectionChanged({
        if ($cmbBaseSheet.SelectedItem -and (Test-Path $txtBasePath.Text)) {
            $sName = $cmbBaseSheet.SelectedItem.ToString()
            if ($script:AppConfig.RememberBasePath -ne $false) {
                $script:AppConfig.BaseSheet = $sName
                Save-AppConfig -Config $script:AppConfig
            }
            try {
                $script:BaseHeaders = [FastExcelHelper]::GetHeaders($txtBasePath.Text, $sName)
                $lbJoinBase.Items.Clear()
                foreach ($h in $script:BaseHeaders) { [void]$lbJoinBase.Items.Add($h) }
                & $CheckProfileAutoMatch
                & $script:UpdateDataMappingPreview $true
            } catch {
                $txtStatusMsg.Text = (Get-UiString 'ErrReadBaseHeaders') -f $_.Exception.Message
            }
        }
    })

    $cmbIncomingSheet.add_SelectionChanged({
        if ($cmbIncomingSheet.SelectedItem -and (Test-Path $txtIncomingPath.Text)) {
            $sName = $cmbIncomingSheet.SelectedItem.ToString()
            try {
                $script:IncomingHeaders = [FastExcelHelper]::GetHeaders($txtIncomingPath.Text, $sName)
                $lbJoinIncoming.Items.Clear()
                foreach ($h in $script:IncomingHeaders) { [void]$lbJoinIncoming.Items.Add($h) }
                & $CheckProfileAutoMatch
                & $script:UpdateDataMappingPreview $true
            } catch {
                $txtStatusMsg.Text = (Get-UiString 'ErrReadIncomingHeaders') -f $_.Exception.Message
            }
        }
    })

    # Browse Base / Incoming
    $btnBrowseBase.add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = (Get-UiString 'FilterExcelCsv')
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            & $LoadBaseFile $dlg.FileName
        }
    })

    $btnBrowseIncoming.add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = (Get-UiString 'FilterExcelCsv')
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            & $LoadIncomingFile $dlg.FileName
        }
    })

    # Universal Drag & Drop Helper for Cards & TextBoxes
    $AttachFileDragDrop = {
        param($TargetElement, [scriptblock]$OnFileLoaded)
        if (-not $TargetElement) { return }
        $TargetElement.AllowDrop = $true
        $TargetElement.add_PreviewDragOver({
            param($s, $e)
            if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
                $e.Effects = [System.Windows.DragDropEffects]::Copy
                $e.Handled = $true
            }
        })
        $TargetElement.add_DragEnter({
            param($s, $e)
            if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
                $TargetElement.BorderBrush = $window.Resources['AccentBlue']
            }
        })
        $TargetElement.add_DragLeave({
            param($s, $e)
            $TargetElement.BorderBrush = $window.Resources['BorderCard']
        })
        $TargetElement.add_Drop({
            param($s, $e)
            $TargetElement.BorderBrush = $window.Resources['BorderCard']
            if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
                $files = $e.Data.GetData([System.Windows.DataFormats]::FileDrop)
                if ($files -and $files.Length -gt 0) {
                    $droppedPath = $files[0]
                    if (Test-Path $droppedPath) {
                        & $OnFileLoaded $droppedPath
                    }
                }
            }
        })
    }

    & $AttachFileDragDrop $cardBaseFile $LoadBaseFile
    & $AttachFileDragDrop $txtBasePath $LoadBaseFile
    & $AttachFileDragDrop $cardIncomingFile $LoadIncomingFile
    & $AttachFileDragDrop $txtIncomingPath $LoadIncomingFile

    # Direct typing / paste support in path textboxes
    $txtBasePath.add_LostFocus({
        if ($txtBasePath.Text -and $txtBasePath.Text.Trim() -ne $script:LoadedBasePath) {
            & $LoadBaseFile $txtBasePath.Text
        }
    })
    $txtBasePath.add_KeyDown({
        param($s, $e)
        if ($e.Key -eq 'Enter') {
            & $LoadBaseFile $txtBasePath.Text
            $e.Handled = $true
        }
    })

    $txtIncomingPath.add_LostFocus({
        if ($txtIncomingPath.Text -and $txtIncomingPath.Text.Trim() -ne $script:LoadedIncomingPath) {
            & $LoadIncomingFile $txtIncomingPath.Text
        }
    })
    $txtIncomingPath.add_KeyDown({
        param($s, $e)
        if ($e.Key -eq 'Enter') {
            & $LoadIncomingFile $txtIncomingPath.Text
            $e.Handled = $true
        }
    })

    # Auto-Map Button
    $btnAutoMap.add_Click({
        if (-not $script:BaseHeaders -or $script:BaseHeaders.Count -eq 0 -or -not $script:IncomingHeaders -or $script:IncomingHeaders.Count -eq 0) {
            $msg = Get-UiString 'ErrLoadFilesBeforeRule'
            $txtStatusMsg.Text = $msg
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $script:MappingRules.Clear()
        $autoRules = Invoke-AutoMapRules -BaseHeaders $script:BaseHeaders -IncomingHeaders $script:IncomingHeaders
        foreach ($r in $autoRules) {
            $script:MappingRules.Add($r)
        }

        # Auto-select Join Key: Only select if a matching key pair exists in BOTH files!
        $lbJoinBase.SelectedItems.Clear()
        $lbJoinIncoming.SelectedItems.Clear()

        $foundJoinKey = $false
        # 1. First priority: Shared ID/Kod/PESEL/Nr column present in both
        foreach ($ih in $script:IncomingHeaders) {
            if ($ih -match '(?i)(id|kod|pesel|nr)') {
                $normIh = $ih.Trim().ToLowerInvariant() -replace '[_\-\s]+', ''
                $matchBase = $script:BaseHeaders | Where-Object {
                    $normBh = $_.Trim().ToLowerInvariant() -replace '[_\-\s]+', ''
                    $normBh -eq $normIh
                } | Select-Object -First 1
                if ($matchBase -and $lbJoinBase.Items.Contains($matchBase) -and $lbJoinIncoming.Items.Contains($ih)) {
                    [void]$lbJoinBase.SelectedItems.Add($matchBase)
                    [void]$lbJoinIncoming.SelectedItems.Add($ih)
                    $foundJoinKey = $true
                    break
                }
            }
        }

        # 2. Second priority: Shared Person/Entity/Name identifier present in both
        if (-not $foundJoinKey) {
            foreach ($ih in $script:IncomingHeaders) {
                if ($ih -match '(?i)(nazwisk|imie|imię|dzieck|klient|uczen|uczeń|pracownik|nazwa|name|surname|firstname|lastname|child|student|employee|client|customer|nachname|vorname|kind|schueler|mitarbeiter|kunde)') {
                    $normIh = $ih.Trim().ToLowerInvariant() -replace '[_\-\s]+', ''
                    $matchBase = $script:BaseHeaders | Where-Object {
                        $normBh = $_.Trim().ToLowerInvariant() -replace '[_\-\s]+', ''
                        $normBh -eq $normIh
                    } | Select-Object -First 1
                    if ($matchBase -and $lbJoinBase.Items.Contains($matchBase) -and $lbJoinIncoming.Items.Contains($ih)) {
                        [void]$lbJoinBase.SelectedItems.Add($matchBase)
                        [void]$lbJoinIncoming.SelectedItems.Add($ih)
                        $foundJoinKey = $true
                        break
                    }
                }
            }
        }

        # 3. Third priority: First column matching by name in both files
        if (-not $foundJoinKey -and $script:IncomingHeaders -and $script:IncomingHeaders.Count -gt 0) {
            $firstIh = $script:IncomingHeaders[0]
            $normFirstIh = $firstIh.Trim().ToLowerInvariant() -replace '[_\-\s]+', ''
            $matchBase = $script:BaseHeaders | Where-Object {
                $normBh = $_.Trim().ToLowerInvariant() -replace '[_\-\s]+', ''
                $normBh -eq $normFirstIh
            } | Select-Object -First 1
            if ($matchBase -and $lbJoinBase.Items.Contains($matchBase) -and $lbJoinIncoming.Items.Contains($firstIh)) {
                [void]$lbJoinBase.SelectedItems.Add($matchBase)
                [void]$lbJoinIncoming.SelectedItems.Add($firstIh)
            }
        }

        if ($script:MappingRules.Count -gt 0) {
            $msg = (Get-UiString 'MsgAutoMapSuccess') -f $script:MappingRules.Count
            $txtStatusMsg.Text = $msg
        } else {
            $msg = Get-UiString 'ErrAutoMapNoMatch'
            $txtStatusMsg.Text = $msg
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'InfoTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        & $script:UpdateDataMappingPreview
    })

    # Rule Editor Dialog (Supports Multi-Select ListBoxes, Concatenate with Separator, and Live Preview)
    $ShowRuleDialog = {
        param([object]$ExistingRule = $null)

        if (-not $script:BaseHeaders -or $script:BaseHeaders.Count -eq 0 -or -not $script:IncomingHeaders -or $script:IncomingHeaders.Count -eq 0) {
            $msg = Get-UiString 'ErrLoadFilesBeforeRule'
            $txtStatusMsg.Text = $msg
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $p = Get-UpdaterThemePalette $script:CurrentTheme
        $ruleWin = New-Object System.Windows.Window -Property @{
            Title                 = if ($ExistingRule) { Get-UiString 'EditRuleTitle' } else { Get-UiString 'AddRuleTitle' }
            Width                 = 620
            Height                = 560
            MinWidth              = 520
            MinHeight             = 480
            WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner
            Owner                 = $window
            Background            = $script:BrushConverter.ConvertFromString($p.BgCard)
            Foreground            = $script:BrushConverter.ConvertFromString($p.TextPrimary)
            FontFamily            = New-Object System.Windows.Media.FontFamily('Segoe UI')
            FontSize              = 13
        }
        $ruleHwnd = (New-Object System.Windows.Interop.WindowInteropHelper($ruleWin)).EnsureHandle()
        Set-WindowDwmTheme -Hwnd $ruleHwnd -IsDark $p.IsDark
        if ($null -ne $window.Resources) {
            foreach ($k in $window.Resources.Keys) { if ($null -ne $window.Resources[$k]) { $ruleWin.Resources[$k] = $window.Resources[$k] } }
        }

        $mainGrid = New-Object System.Windows.Controls.Grid -Property @{ Margin = New-Object System.Windows.Thickness(18) }
        $mainGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto }))
        $mainGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) }))
        $mainGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto }))
        $mainGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto }))
        $mainGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto }))

        # Row 0: Hint banner
        $txtHint = New-Object System.Windows.Controls.TextBlock -Property @{
            Text         = Get-UiString 'RuleMultiSelectHint'
            Foreground   = $script:BrushConverter.ConvertFromString($p.TextSecondary)
            FontSize     = 11
            Margin       = New-Object System.Windows.Thickness(0, 0, 0, 10)
            TextWrapping = [System.Windows.TextWrapping]::Wrap
        }
        [System.Windows.Controls.Grid]::SetRow($txtHint, 0)
        [void]$mainGrid.Children.Add($txtHint)

        # Row 1: Columns ListBoxes Grid
        $colsGrid = New-Object System.Windows.Controls.Grid
        $colsGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) }))
        $colsGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(14, [System.Windows.GridUnitType]::Pixel) }))
        $colsGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) }))

        # Left: Base columns
        $spBase = New-Object System.Windows.Controls.Grid
        $spBase.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto }))
        $spBase.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) }))
        $lblB = New-Object System.Windows.Controls.TextBlock -Property @{
            Text       = Get-UiString 'ColBase'
            Foreground = $script:BrushConverter.ConvertFromString($p.TextPrimary)
            FontWeight = [System.Windows.FontWeights]::SemiBold
            Margin     = New-Object System.Windows.Thickness(0, 0, 0, 6)
        }
        $lbBase = New-Object System.Windows.Controls.ListBox -Property @{
            Background    = $script:BrushConverter.ConvertFromString($p.BgInput)
            Foreground    = $script:BrushConverter.ConvertFromString($p.TextPrimary)
            BorderBrush   = $script:BrushConverter.ConvertFromString($p.BorderInput)
            SelectionMode = [System.Windows.Controls.SelectionMode]::Extended
        }
        foreach ($h in $script:BaseHeaders) { [void]$lbBase.Items.Add($h) }
        [System.Windows.Controls.Grid]::SetRow($lblB, 0)
        [System.Windows.Controls.Grid]::SetRow($lbBase, 1)
        [void]$spBase.Children.Add($lblB)
        [void]$spBase.Children.Add($lbBase)
        [System.Windows.Controls.Grid]::SetColumn($spBase, 0)
        [void]$colsGrid.Children.Add($spBase)

        # Right: Incoming columns
        $spUpd = New-Object System.Windows.Controls.Grid
        $spUpd.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto }))
        $spUpd.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) }))
        $lblU = New-Object System.Windows.Controls.TextBlock -Property @{
            Text       = Get-UiString 'ColIncoming'
            Foreground = $script:BrushConverter.ConvertFromString($p.TextPrimary)
            FontWeight = [System.Windows.FontWeights]::SemiBold
            Margin     = New-Object System.Windows.Thickness(0, 0, 0, 6)
        }
        $lbUpd = New-Object System.Windows.Controls.ListBox -Property @{
            Background    = $script:BrushConverter.ConvertFromString($p.BgInput)
            Foreground    = $script:BrushConverter.ConvertFromString($p.TextPrimary)
            BorderBrush   = $script:BrushConverter.ConvertFromString($p.BorderInput)
            SelectionMode = [System.Windows.Controls.SelectionMode]::Extended
        }
        foreach ($h in $script:IncomingHeaders) { [void]$lbUpd.Items.Add($h) }
        [System.Windows.Controls.Grid]::SetRow($lblU, 0)
        [System.Windows.Controls.Grid]::SetRow($lbUpd, 1)
        [void]$spUpd.Children.Add($lblU)
        [void]$spUpd.Children.Add($lbUpd)
        [System.Windows.Controls.Grid]::SetColumn($spUpd, 2)
        [void]$colsGrid.Children.Add($spUpd)

        [System.Windows.Controls.Grid]::SetRow($colsGrid, 1)
        [void]$mainGrid.Children.Add($colsGrid)

        # Row 2: Merge Mode, Separator, and Diff Policy controls
        $modeGrid = New-Object System.Windows.Controls.Grid -Property @{ Margin = New-Object System.Windows.Thickness(0, 10, 0, 10) }
        $modeGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto }))
        $modeGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto }))
        $modeGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) }))
        $modeGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(14, [System.Windows.GridUnitType]::Pixel) }))
        $modeGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) }))

        $spMode = New-Object System.Windows.Controls.StackPanel
        $lblM = New-Object System.Windows.Controls.TextBlock -Property @{
            Text       = Get-UiString 'ColMergeMode'
            Foreground = $script:BrushConverter.ConvertFromString($p.TextSecondary)
            Margin     = New-Object System.Windows.Thickness(0, 0, 0, 4)
        }
        $cmbM = New-Object System.Windows.Controls.ComboBox -Property @{
            Background  = $script:BrushConverter.ConvertFromString($p.BgInput)
            Foreground  = $script:BrushConverter.ConvertFromString($p.TextPrimary)
            BorderBrush = $script:BrushConverter.ConvertFromString($p.BorderInput)
        }
        [void]$cmbM.Items.Add('Exact'); [void]$cmbM.Items.Add('Concatenate'); [void]$cmbM.Items.Add('FirstNonEmpty')
        $cmbM.SelectedIndex = 0
        [void]$spMode.Children.Add($lblM); [void]$spMode.Children.Add($cmbM)
        [System.Windows.Controls.Grid]::SetRow($spMode, 0)
        [System.Windows.Controls.Grid]::SetColumn($spMode, 0)
        [void]$modeGrid.Children.Add($spMode)

        $spSep = New-Object System.Windows.Controls.StackPanel
        $lblS = New-Object System.Windows.Controls.TextBlock -Property @{
            Text       = Get-UiString 'ColSeparator'
            Foreground = $script:BrushConverter.ConvertFromString($p.TextSecondary)
            Margin     = New-Object System.Windows.Thickness(0, 0, 0, 4)
        }
        $txtS = New-Object System.Windows.Controls.TextBox -Property @{
            Background  = $script:BrushConverter.ConvertFromString($p.BgInput)
            Foreground  = $script:BrushConverter.ConvertFromString($p.TextPrimary)
            BorderBrush = $script:BrushConverter.ConvertFromString($p.BorderInput)
            Padding     = New-Object System.Windows.Thickness(6, 4, 6, 4)
            Text        = ', '
        }
        [void]$spSep.Children.Add($lblS); [void]$spSep.Children.Add($txtS)
        [System.Windows.Controls.Grid]::SetRow($spSep, 0)
        [System.Windows.Controls.Grid]::SetColumn($spSep, 2)
        [void]$modeGrid.Children.Add($spSep)

        # Row 1: Diff Policy ComboBox
        $spDiff = New-Object System.Windows.Controls.StackPanel -Property @{ Margin = New-Object System.Windows.Thickness(0, 8, 0, 0) }
        $lblD = New-Object System.Windows.Controls.TextBlock -Property @{
            Text       = Get-UiString 'ColDiffPolicy'
            Foreground = $script:BrushConverter.ConvertFromString($p.TextSecondary)
            Margin     = New-Object System.Windows.Thickness(0, 0, 0, 4)
        }
        $cmbDiff = New-Object System.Windows.Controls.ComboBox -Property @{
            Background  = $script:BrushConverter.ConvertFromString($p.BgInput)
            Foreground  = $script:BrushConverter.ConvertFromString($p.TextPrimary)
            BorderBrush = $script:BrushConverter.ConvertFromString($p.BorderInput)
        }
        $diffPolicies = @(
            @{ Tag = 'TrackChanges';        Text = (Get-UiString 'DiffPolicyTrackChanges' 'Śledź zmiany (standard)') },
            @{ Tag = 'IgnoreChanges';       Text = (Get-UiString 'DiffPolicyIgnoreChanges' 'Ignoruj zmiany (tylko nowe wpisy)') },
            @{ Tag = 'NormalizePostalCode'; Text = (Get-UiString 'DiffPolicyNormalizePostal' 'Ignoruj obecność kodu pocztowego') },
            @{ Tag = 'FuzzyContainment';    Text = (Get-UiString 'DiffPolicyFuzzyContainment' 'Ignoruj dopiski w nazwach (zawieranie tekstu)') }
        )
        foreach ($dpItem in $diffPolicies) {
            $cbi = New-Object System.Windows.Controls.ComboBoxItem
            $cbi.Tag = $dpItem.Tag
            $cbi.Content = $dpItem.Text
            [void]$cmbDiff.Items.Add($cbi)
        }
        $cmbDiff.SelectedIndex = 0
        [void]$spDiff.Children.Add($lblD); [void]$spDiff.Children.Add($cmbDiff)
        [System.Windows.Controls.Grid]::SetRow($spDiff, 1)
        [System.Windows.Controls.Grid]::SetColumn($spDiff, 0)
        [System.Windows.Controls.Grid]::SetColumnSpan($spDiff, 3)
        [void]$modeGrid.Children.Add($spDiff)

        [System.Windows.Controls.Grid]::SetRow($modeGrid, 2)
        [void]$mainGrid.Children.Add($modeGrid)

        # Row 3: Live Preview Box
        $borderPreview = New-Object System.Windows.Controls.Border -Property @{
            Background      = $script:BrushConverter.ConvertFromString($p.BgInput)
            BorderBrush     = $script:BrushConverter.ConvertFromString($p.BorderCard)
            BorderThickness = [System.Windows.Thickness]::new(1)
            CornerRadius    = [System.Windows.CornerRadius]::new(6)
            Padding         = New-Object System.Windows.Thickness(10)
            Margin          = New-Object System.Windows.Thickness(0, 0, 0, 14)
        }
        $spPreviewContent = New-Object System.Windows.Controls.StackPanel
        $lblPreviewTitle = New-Object System.Windows.Controls.TextBlock -Property @{
            Text       = Get-UiString 'RulePreviewTitle'
            Foreground = $script:BrushConverter.ConvertFromString($p.TextSecondary)
            FontSize   = 11
            FontWeight = [System.Windows.FontWeights]::SemiBold
            Margin     = New-Object System.Windows.Thickness(0, 0, 0, 4)
        }
        $txtPreviewText = New-Object System.Windows.Controls.TextBlock -Property @{
            Foreground   = $script:BrushConverter.ConvertFromString($p.TextPrimary)
            FontFamily   = New-Object System.Windows.Media.FontFamily('Consolas, Segoe UI')
            FontSize     = 12
            TextWrapping = [System.Windows.TextWrapping]::Wrap
        }
        [void]$spPreviewContent.Children.Add($lblPreviewTitle)
        [void]$spPreviewContent.Children.Add($txtPreviewText)
        $borderPreview.Child = $spPreviewContent

        [System.Windows.Controls.Grid]::SetRow($borderPreview, 3)
        [void]$mainGrid.Children.Add($borderPreview)

        # Row 4: Buttons
        $btnSp = New-Object System.Windows.Controls.StackPanel -Property @{
            Orientation         = [System.Windows.Controls.Orientation]::Horizontal
            HorizontalAlignment = [System.Windows.HorizontalAlignment]::Right
        }
        $btnSave = New-Object System.Windows.Controls.Button -Property @{
            Content    = Get-UiString 'SettingsBtnSave'
            Background = $script:BrushConverter.ConvertFromString($p.AccentBlue)
            Foreground = $script:BrushConverter.ConvertFromString('#FFFFFF')
            FontWeight = [System.Windows.FontWeights]::Bold
            Padding    = New-Object System.Windows.Thickness(16, 6, 16, 6)
            Margin     = New-Object System.Windows.Thickness(0, 0, 8, 0)
            Cursor     = [System.Windows.Input.Cursors]::Hand
        }
        $btnCancel = New-Object System.Windows.Controls.Button -Property @{
            Content         = Get-UiString 'SettingsBtnCancel'
            Background      = $script:BrushConverter.ConvertFromString($p.BtnSecondaryBg)
            Foreground      = $script:BrushConverter.ConvertFromString($p.BtnSecondaryFg)
            BorderBrush     = $script:BrushConverter.ConvertFromString($p.BorderCard)
            BorderThickness = [System.Windows.Thickness]::new(1)
            Padding         = New-Object System.Windows.Thickness(14, 6, 14, 6)
            Cursor          = [System.Windows.Input.Cursors]::Hand
        }
        [void]$btnSp.Children.Add($btnSave); [void]$btnSp.Children.Add($btnCancel)
        [System.Windows.Controls.Grid]::SetRow($btnSp, 4)
        [void]$mainGrid.Children.Add($btnSp)

        # Update Preview handler
        $UpdatePreview = {
            $bSel = @($lbBase.SelectedItems)
            $uSel = @($lbUpd.SelectedItems)
            $mSel = if ($cmbM.SelectedItem) { $cmbM.SelectedItem.ToString() } else { 'Exact' }
            $sVal = $txtS.Text

            $bStr = if ($bSel.Count -gt 0) {
                if ($mSel -eq 'Concatenate' -and $sVal) { $bSel -join " $($sVal.Trim()) " } else { $bSel -join ', ' }
            } else { '(brak / none)' }

            $uStr = if ($uSel.Count -gt 0) {
                if ($mSel -eq 'Concatenate' -and $sVal) { $uSel -join " $($sVal.Trim()) " } else { $uSel -join ', ' }
            } else { '(brak / none)' }

            $modeExtra = if ($mSel -eq 'Concatenate') { " (Separator: '$sVal')" } else { "" }
            $dpSel = if ($cmbDiff.SelectedItem -is [System.Windows.Controls.ComboBoxItem]) { $cmbDiff.SelectedItem.Content.ToString() } else { 'TrackChanges' }
            $txtPreviewText.Text = "Base:   $bStr`nUpdate: $uStr`nMode:   $mSel$modeExtra`nPolicy: $dpSel"
        }

        $lbBase.add_SelectionChanged({ & $UpdatePreview })
        $lbUpd.add_SelectionChanged({ & $UpdatePreview })
        $cmbM.add_SelectionChanged({
            $isConcat = ($cmbM.SelectedItem -and $cmbM.SelectedItem.ToString() -eq 'Concatenate')
            if ($isConcat -and [string]::IsNullOrEmpty($txtS.Text)) { $txtS.Text = ', ' }
            & $UpdatePreview
        })
        $cmbDiff.add_SelectionChanged({ & $UpdatePreview })
        $txtS.add_TextChanged({ & $UpdatePreview })

        # Pre-populate if ExistingRule
        if ($ExistingRule) {
            $exBCols = if ($ExistingRule.BaseColumns) { @($ExistingRule.BaseColumns) } else { @() }
            $exUCols = if ($ExistingRule.UpdateColumns) { @($ExistingRule.UpdateColumns) } else { @() }
            foreach ($c in $exBCols) {
                if ($lbBase.Items.Contains($c)) { [void]$lbBase.SelectedItems.Add($c) }
            }
            foreach ($c in $exUCols) {
                if ($lbUpd.Items.Contains($c)) { [void]$lbUpd.SelectedItems.Add($c) }
            }
            $mi = @('Exact', 'Concatenate', 'FirstNonEmpty').IndexOf($ExistingRule.MergeMode)
            if ($mi -ge 0) { $cmbM.SelectedIndex = $mi }
            $txtS.Text = if ($null -ne $ExistingRule.Separator) { $ExistingRule.Separator } else { '' }

            $curDp = if ($ExistingRule.DiffPolicy) { $ExistingRule.DiffPolicy } else { 'TrackChanges' }
            for ($i = 0; $i -lt $cmbDiff.Items.Count; $i++) {
                if ($cmbDiff.Items[$i].Tag -eq $curDp) {
                    $cmbDiff.SelectedIndex = $i
                    break
                }
            }
        } else {
            if ($lbBase.Items.Count -gt 0) { $lbBase.SelectedIndex = 0 }
            if ($lbUpd.Items.Count -gt 0) { $lbUpd.SelectedIndex = 0 }
        }
        & $UpdatePreview

        # Save action
        $btnSave.add_Click({
            $bSel = @($lbBase.SelectedItems)
            $uSel = @($lbUpd.SelectedItems)
            if ($bSel.Count -eq 0 -or $uSel.Count -eq 0) {
                $msg = Get-UiString 'ErrSelectAtLeastOneCol'
                [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }

            $bColsArr = [string[]]@($bSel | ForEach-Object { $_.ToString() })
            $uColsArr = [string[]]@($uSel | ForEach-Object { $_.ToString() })
            $mModeVal = if ($cmbM.SelectedItem) { $cmbM.SelectedItem.ToString() } else { 'Exact' }
            $sepVal   = $txtS.Text
            $dpVal    = if ($cmbDiff.SelectedItem -is [System.Windows.Controls.ComboBoxItem]) { $cmbDiff.SelectedItem.Tag.ToString() } elseif ($cmbDiff.SelectedItem) { $cmbDiff.SelectedItem.ToString() } else { 'TrackChanges' }
            $dpDisplay = Get-DiffPolicyDisplay $dpVal

            if ($ExistingRule) {
                $ExistingRule.BaseColumns       = $bColsArr
                $ExistingRule.UpdateColumns     = $uColsArr
                $ExistingRule.MergeMode         = $mModeVal
                $ExistingRule.Separator         = $sepVal
                $ExistingRule.DiffPolicy        = $dpVal
                $ExistingRule.DiffPolicyDisplay = $dpDisplay
                $ExistingRule.BaseColsStr       = ($bColsArr -join ', ')
                $ExistingRule.UpdColsStr        = ($uColsArr -join ', ')
                $gridMappingRules.Items.Refresh()
            } else {
                $script:MappingRules.Add([PSCustomObject]@{
                    BaseColumns       = $bColsArr
                    UpdateColumns     = $uColsArr
                    MergeMode         = $mModeVal
                    Separator         = $sepVal
                    DiffPolicy        = $dpVal
                    DiffPolicyDisplay = $dpDisplay
                    BaseColsStr       = ($bColsArr -join ', ')
                    UpdColsStr        = ($uColsArr -join ', ')
                })
            }
            $ruleWin.Close()
            & $script:UpdateDataMappingPreview
        })

        $btnCancel.add_Click({ $ruleWin.Close() })

        $ruleWin.Content = $mainGrid
        [void]$ruleWin.ShowDialog()
    }

    # Add Rule Button
    $btnAddRule.add_Click({
        & $ShowRuleDialog
    })

    # Edit Rule Button
    if ($btnEditRule) {
        $btnEditRule.add_Click({
            $sel = if ($gridMappingRules.SelectedItem) { $gridMappingRules.SelectedItem } else { $null }
            if ($sel) {
                & $ShowRuleDialog -ExistingRule $sel
            } else {
                $msg = Get-UiString 'ErrSelectRuleToEdit'
                [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
        })
    }

    # Double-click on Mapping Rules DataGrid to Edit Rule
    $gridMappingRules.add_MouseDoubleClick({
        if ($gridMappingRules.SelectedItem) {
            & $ShowRuleDialog -ExistingRule $gridMappingRules.SelectedItem
        }
    })

    # Remove Rule Button
    $btnRemoveRule.add_Click({
        $selectedList = @($gridMappingRules.SelectedItems)
        if ($selectedList.Count -eq 0 -and $gridMappingRules.SelectedItem) {
            $selectedList = @($gridMappingRules.SelectedItem)
        }
        if ($selectedList.Count -gt 0) {
            foreach ($item in $selectedList) {
                if ($script:MappingRules.Contains($item)) {
                    [void]$script:MappingRules.Remove($item)
                }
            }
            & $script:UpdateDataMappingPreview
        }
    })

    # Preview sample row navigation & refresh buttons
    $btnPrevSampleRow.add_Click({
        if ($script:PreviewRowIndex -gt 0) {
            $script:PreviewRowIndex--
            & $script:UpdateDataMappingPreview
        }
    })

    $btnNextSampleRow.add_Click({
        $maxCount = 0
        if ($script:CachedIncomingSampleRows) { $maxCount = [Math]::Max($maxCount, $script:CachedIncomingSampleRows.Count) }
        if ($script:CachedBaseSampleRows) { $maxCount = [Math]::Max($maxCount, $script:CachedBaseSampleRows.Count) }
        if ($maxCount -gt 0 -and $script:PreviewRowIndex -lt ($maxCount - 1)) {
            $script:PreviewRowIndex++
            & $script:UpdateDataMappingPreview
        }
    })

    $btnRefreshPreview.add_Click({
        & $script:UpdateDataMappingPreview $true
    })

    # Mapping Rules Live Search / Filter
    $script:ApplyMappingFilter = {
        if (-not $gridMappingRules -or -not $gridMappingRules.ItemsSource) { return }
        $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($gridMappingRules.ItemsSource)
        if (-not $view) { return }
        $q = if ($txtSearchMapping -and $txtSearchMapping.Text) { $txtSearchMapping.Text.Trim().ToLowerInvariant() } else { '' }
        if ([string]::IsNullOrEmpty($q)) {
            $view.Filter = $null
        } else {
            $view.Filter = [System.Predicate[object]]{
                param($it)
                if (-not $it) { return $false }
                $b = if ($it.BaseColsStr) { $it.BaseColsStr.ToLowerInvariant() } elseif ($it.BaseColumns) { ([string]::Join(' ', $it.BaseColumns)).ToLowerInvariant() } else { '' }
                $u = if ($it.UpdColsStr) { $it.UpdColsStr.ToLowerInvariant() } elseif ($it.UpdateColumns) { ([string]::Join(' ', $it.UpdateColumns)).ToLowerInvariant() } else { '' }
                $m = if ($it.MergeMode) { $it.MergeMode.ToLowerInvariant() } else { '' }
                return ($b.Contains($q) -or $u.Contains($q) -or $m.Contains($q))
            }
        }
        $view.Refresh()
    }
    if ($txtSearchMapping) {
        $txtSearchMapping.add_TextChanged({
            if ($btnSearchMappingClear) {
                $btnSearchMappingClear.Visibility = if ([string]::IsNullOrEmpty($txtSearchMapping.Text)) {
                    [System.Windows.Visibility]::Collapsed
                } else {
                    [System.Windows.Visibility]::Visible
                }
            }
            if ($script:ApplyMappingFilter) { & $script:ApplyMappingFilter }
        })
    }
    if ($btnSearchMappingClear) {
        $btnSearchMappingClear.add_Click({
            if ($txtSearchMapping) {
                $txtSearchMapping.Text = ''
                $txtSearchMapping.Focus()
            }
        })
    }

    # Live Preview Immediate Reaction to Join Key selection changes
    if ($lbJoinBase) {
        $lbJoinBase.add_SelectionChanged({
            & $script:UpdateDataMappingPreview
        })
    }
    if ($lbJoinIncoming) {
        $lbJoinIncoming.add_SelectionChanged({
            & $script:UpdateDataMappingPreview
        })
    }

    # Save Profile Button
    $btnSaveProfile.add_Click({
        if ($script:MappingRules.Count -eq 0) {
            $msg = Get-UiString 'ErrDefineRules'
            $txtStatusMsg.Text = $msg
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $baseKeys = @($lbJoinBase.SelectedItems | ForEach-Object { $_.ToString() })
        $incKeys  = @($lbJoinIncoming.SelectedItems | ForEach-Object { $_.ToString() })
        if ($baseKeys.Length -eq 0 -or $incKeys.Length -eq 0) {
            $msg = Get-UiString 'ErrSelectJoinKeys'
            $txtStatusMsg.Text = $msg
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        if ($baseKeys.Length -ne $incKeys.Length) {
            $msg = (Get-UiString 'ErrJoinKeyCountMismatch') -f $baseKeys.Length, ($baseKeys -join ', '), $incKeys.Length, ($incKeys -join ', ')
            $txtStatusMsg.Text = $msg
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $defName = [System.IO.Path]::GetFileNameWithoutExtension($txtIncomingPath.Text)
        if ([string]::IsNullOrWhiteSpace($defName)) { $defName = 'Profil_' + (Get-Date -Format 'yyyyMMdd_HHmm') }

        $sheetName = if ($cmbIncomingSheet.SelectedItem) { $cmbIncomingSheet.SelectedItem.ToString() } else { '' }
        $fp = Compute-HeaderFingerprint -Headers $script:IncomingHeaders -SheetName $sheetName

        $rulesList = [System.Collections.Generic.List[object]]::new()
        foreach ($r in $script:MappingRules) {
            $rulesList.Add(@{
                BaseColumns   = @($r.BaseColumns)
                UpdateColumns = @($r.UpdateColumns)
                MergeMode     = $r.MergeMode
                Separator     = $r.Separator
                DiffPolicy    = if ($r.PSObject.Properties['DiffPolicy'] -and $r.DiffPolicy) { $r.DiffPolicy } else { 'TrackChanges' }
            })
        }

        $compOptions = @{
            IgnoreCase         = if ($script:chkIgnoreCase) { [bool]$script:chkIgnoreCase.IsChecked } else { $true }
            Trim               = if ($script:chkTrimWhitespace) { [bool]$script:chkTrimWhitespace.IsChecked } else { $true }
            TrimWhitespace     = if ($script:chkTrimWhitespace) { [bool]$script:chkTrimWhitespace.IsChecked } else { $true }
            IgnoreSpecialChars = if ($script:chkIgnoreSpecialChars) { [bool]$script:chkIgnoreSpecialChars.IsChecked } else { $true }
            IgnoreAllSpaces    = if ($script:chkIgnoreAllSpaces) { [bool]$script:chkIgnoreAllSpaces.IsChecked } else { $true }
        }

        $prof = [ordered]@{
            SchemaVersion     = '2.0'
            Name              = $defName
            CreatedUtc        = (Get-Date).ToUniversalTime().ToString('o')
            HeaderFingerprint = $fp
            BaseFilePath      = $txtBasePath.Text
            BaseSheet         = if ($cmbBaseSheet.SelectedItem) { $cmbBaseSheet.SelectedItem.ToString() } else { '' }
            SourceSheetName   = $sheetName
            JoinKeyBase       = $baseKeys
            JoinKeyUpdate     = $incKeys
            MappingRules      = $rulesList
            CompareOptions    = $compOptions
        }

        [void](Save-MappingProfile -Profile $prof -StorePath $script:AppConfig.ProfileStorePath)
        $badgeProfile.Visibility = [System.Windows.Visibility]::Visible
        $txtProfileBadge.Text = "$((Get-UiString 'BadgeProfileManual')): $defName"
        [System.Windows.Forms.MessageBox]::Show(((Get-UiString 'ProfileSavedMsg') -f $defName), (Get-UiString 'ProfileSavedTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })

    # Render Detail Pane
    $script:RenderDetailPane = $RenderDetailPane = {
        param($Item, [switch]$SkipBaseRowRefresh)
        if (-not $panelDiffContainer) {
            $panelDiffContainer = if ($window) { $window.FindName('panelDiffContainer') } else { $script:ActiveWindow.FindName('panelDiffContainer') }
        }
        if (-not $txtDetailHeader) {
            $txtDetailHeader = if ($window) { $window.FindName('txtDetailHeader') } else { $script:ActiveWindow.FindName('txtDetailHeader') }
        }
        if (-not $panelDiffContainer) { return }
        $panelDiffContainer.Children.Clear()
        if (-not $Item) {
            if ($txtDetailHeader) { $txtDetailHeader.Text = Get-UiString 'SelectRowHeader' }
            return
        }

        $p = Get-UpdaterThemePalette $script:CurrentTheme
        if ($Item -and -not $Item.SelectedCells) { $Item.SelectedCells = @{} }
        $rec = $Item.Record
        $status = $rec.Status

        if ($status -eq 'New') {
            $txtDetailHeader.Text = "$((Get-UiString 'StatusBadgeNew')): $($Item.Title)"
            $grid = New-Object System.Windows.Controls.Grid
            $grid.Margin = New-Object System.Windows.Thickness(0, 8, 0, 0)
            [void]$grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(180) }))
            [void]$grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) }))

            $rowIdx = 0
            foreach ($k in $rec.ProjectedRow.Keys) {
                [void]$grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto }))
                $lbl = New-Object System.Windows.Controls.TextBlock -Property @{
                    Text = $k
                    Foreground = $script:BrushConverter.ConvertFromString($p.TextSecondary)
                    FontWeight = [System.Windows.FontWeights]::SemiBold
                    Margin = New-Object System.Windows.Thickness(0, 4, 10, 4)
                }
                $val = New-Object System.Windows.Controls.TextBlock -Property @{
                    Text = $(if ($rec.ProjectedRow[$k]) { $rec.ProjectedRow[$k].ToString() } else { '' })
                    Foreground = $script:BrushConverter.ConvertFromString($p.TextPrimary)
                    Margin = New-Object System.Windows.Thickness(0, 4, 0, 4)
                }
                [System.Windows.Controls.Grid]::SetRow($lbl, $rowIdx)
                [System.Windows.Controls.Grid]::SetColumn($lbl, 0)
                [System.Windows.Controls.Grid]::SetRow($val, $rowIdx)
                [System.Windows.Controls.Grid]::SetColumn($val, 1)
                [void]$grid.Children.Add($lbl)
                [void]$grid.Children.Add($val)
                $rowIdx++
            }
            [void]$panelDiffContainer.Children.Add($grid)

        } elseif ($status -eq 'Changed' -or $status -eq 'Removed') {
            $baseRowStr = if ($rec.MatchedBaseRow) { (Get-UiString 'RowNumberFormat') -f $rec.MatchedBaseRow.RowNumber } else { '' }
            $txtDetailHeader.Text = "$((Get-UiString 'StatusBadgeChanged')): $($Item.Title) ($baseRowStr)"

            # Row Action Tools (Select/Deselect All in Row)
            $spRowTools = New-Object System.Windows.Controls.StackPanel -Property @{
                Orientation = [System.Windows.Controls.Orientation]::Horizontal
                Margin      = New-Object System.Windows.Thickness(0, 0, 0, 8)
            }
            $btnSelAllRow = New-Object System.Windows.Controls.Button -Property @{
                Content         = Get-UiString 'BtnSelectAllRowFields' 'Zaznacz wszystkie w wierszu'
                Background      = $script:BrushConverter.ConvertFromString($p.BgCardHover)
                Foreground      = $script:BrushConverter.ConvertFromString($p.TextPrimary)
                BorderBrush     = $script:BrushConverter.ConvertFromString($p.BorderCard)
                BorderThickness = New-Object System.Windows.Thickness(1)
                Padding         = New-Object System.Windows.Thickness(8, 3, 8, 3)
                Margin          = New-Object System.Windows.Thickness(0, 0, 6, 0)
                FontSize        = 11
                Cursor          = [System.Windows.Input.Cursors]::Hand
            }
            $btnDeselAllRow = New-Object System.Windows.Controls.Button -Property @{
                Content         = Get-UiString 'BtnDeselectAllRowFields' 'Odznacz wszystkie w wierszu'
                Background      = $script:BrushConverter.ConvertFromString($p.BgCardHover)
                Foreground      = $script:BrushConverter.ConvertFromString($p.TextPrimary)
                BorderBrush     = $script:BrushConverter.ConvertFromString($p.BorderCard)
                BorderThickness = New-Object System.Windows.Thickness(1)
                Padding         = New-Object System.Windows.Thickness(8, 3, 8, 3)
                FontSize        = 11
                Cursor          = [System.Windows.Input.Cursors]::Hand
            }
            $allRowCheckboxes = [System.Collections.Generic.List[object]]::new()
            $btnSelAllRow.add_Click({
                foreach ($c in $allRowCheckboxes) { $c.IsChecked = $true }
                & $script:UpdateStagingSummary
            })
            $btnDeselAllRow.add_Click({
                foreach ($c in $allRowCheckboxes) { $c.IsChecked = $false }
                & $script:UpdateStagingSummary
            })
            [void]$spRowTools.Children.Add($btnSelAllRow)
            [void]$spRowTools.Children.Add($btnDeselAllRow)
            [void]$panelDiffContainer.Children.Add($spRowTools)

            # Diff Table with per-cell checkboxes (Q8) and inline editing
            $grid = New-Object System.Windows.Controls.Grid
            $grid.Margin = New-Object System.Windows.Thickness(0, 0, 0, 0)
            [void]$grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::Auto }))
            [void]$grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(180) }))
            [void]$grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) }))
            [void]$grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) }))

            # Header Row
            [void]$grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto }))
            $hApp = New-Object System.Windows.Controls.TextBlock -Property @{ Text = Get-UiString 'ColApply'; Foreground = $script:BrushConverter.ConvertFromString($p.TextSecondary); FontWeight = [System.Windows.FontWeights]::Bold; Margin = New-Object System.Windows.Thickness(0, 0, 10, 8) }
            $hFld = New-Object System.Windows.Controls.TextBlock -Property @{ Text = Get-UiString 'ColField'; Foreground = $script:BrushConverter.ConvertFromString($p.TextSecondary); FontWeight = [System.Windows.FontWeights]::Bold; Margin = New-Object System.Windows.Thickness(0, 0, 10, 8) }
            $hCur = New-Object System.Windows.Controls.TextBlock -Property @{ Text = Get-UiString 'ColCurrentBase'; Foreground = $script:BrushConverter.ConvertFromString($p.TextSecondary); FontWeight = [System.Windows.FontWeights]::Bold; Margin = New-Object System.Windows.Thickness(0, 0, 10, 8) }
            $hInc = New-Object System.Windows.Controls.TextBlock -Property @{ Text = Get-UiString 'ColIncomingValue'; Foreground = $script:BrushConverter.ConvertFromString($p.TextSecondary); FontWeight = [System.Windows.FontWeights]::Bold; Margin = New-Object System.Windows.Thickness(0, 0, 0, 8) }

            [System.Windows.Controls.Grid]::SetRow($hApp, 0); [System.Windows.Controls.Grid]::SetColumn($hApp, 0); [void]$grid.Children.Add($hApp)
            [System.Windows.Controls.Grid]::SetRow($hFld, 0); [System.Windows.Controls.Grid]::SetColumn($hFld, 1); [void]$grid.Children.Add($hFld)
            [System.Windows.Controls.Grid]::SetRow($hCur, 0); [System.Windows.Controls.Grid]::SetColumn($hCur, 2); [void]$grid.Children.Add($hCur)
            [System.Windows.Controls.Grid]::SetRow($hInc, 0); [System.Windows.Controls.Grid]::SetColumn($hInc, 3); [void]$grid.Children.Add($hInc)

            $rIdx = 1
            foreach ($chg in $rec.Changes) {
                [void]$grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto }))

                # CheckBox for selective update (Q8)
                $colKey = $chg.BaseColumn
                $chk = New-Object System.Windows.Controls.CheckBox -Property @{
                    IsChecked = if ($Item.SelectedCells) { ($Item.SelectedCells[$colKey] -ne $false) } else { $true }
                    VerticalAlignment = [System.Windows.VerticalAlignment]::Center
                    Margin = New-Object System.Windows.Thickness(0, 4, 10, 4)
                }
                [void]$allRowCheckboxes.Add($chk)
                $chk.add_Checked({
                    $Item.SelectedCells[$colKey] = $true
                    $chg.SelectedForUpdate = $true
                    & $script:UpdateStagingSummary
                })
                $chk.add_Unchecked({
                    $Item.SelectedCells[$colKey] = $false
                    $chg.SelectedForUpdate = $false
                    & $script:UpdateStagingSummary
                })

                $txtCol = New-Object System.Windows.Controls.TextBlock -Property @{
                    Text = $chg.BaseColumn
                    Foreground = $script:BrushConverter.ConvertFromString($p.TextPrimary)
                    FontWeight = [System.Windows.FontWeights]::SemiBold
                    VerticalAlignment = [System.Windows.VerticalAlignment]::Center
                    Margin = New-Object System.Windows.Thickness(0, 4, 10, 4)
                }

                $txtOld = New-Object System.Windows.Controls.TextBlock -Property @{
                    Text = $(if ($chg.OldValue) { $chg.OldValue.ToString() } else { '' })
                    Foreground = $script:BrushConverter.ConvertFromString($p.DiffCellOldFg)
                    VerticalAlignment = [System.Windows.VerticalAlignment]::Center
                    Margin = New-Object System.Windows.Thickness(0, 4, 10, 4)
                }

                $isDirty = ($chg.Contains('CustomEdited') -and $chg.CustomEdited -eq $true)
                $bdrBrushHex = if ($isDirty) { $p.AccentBlue } else { $p.DiffCellNewBorder }
                $bdrThickVal = if ($isDirty) { 2 } else { 1 }
                $bdrNew = New-Object System.Windows.Controls.Border -Property @{
                    Background      = $script:BrushConverter.ConvertFromString($p.DiffCellNewBg)
                    BorderBrush     = $script:BrushConverter.ConvertFromString($bdrBrushHex)
                    BorderThickness = New-Object System.Windows.Thickness($bdrThickVal)
                    CornerRadius    = New-Object System.Windows.CornerRadius(4)
                    Padding         = New-Object System.Windows.Thickness(4, 2, 4, 2)
                    Margin          = New-Object System.Windows.Thickness(0, 2, 0, 2)
                }

                $gridNewCell = New-Object System.Windows.Controls.Grid
                [void]$gridNewCell.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) }))
                [void]$gridNewCell.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::Auto }))

                $txtEditNew = New-Object System.Windows.Controls.TextBox -Property @{
                    Text            = $(if ($null -ne $chg.NewValue) { $chg.NewValue.ToString() } else { '' })
                    Foreground      = $script:BrushConverter.ConvertFromString($p.DiffCellNewFg)
                    Background      = [System.Windows.Media.Brushes]::Transparent
                    BorderThickness = New-Object System.Windows.Thickness(0)
                    FontWeight      = [System.Windows.FontWeights]::Bold
                    VerticalAlignment = [System.Windows.VerticalAlignment]::Center
                    Padding         = New-Object System.Windows.Thickness(2, 1, 2, 1)
                }

                $btnRevert = New-Object System.Windows.Controls.Button -Property @{
                    Content         = '↺'
                    ToolTip         = (Get-UiString 'TooltipRevertOriginal' 'Przywróć wartość z pliku zmian')
                    FontSize        = 12
                    FontWeight      = [System.Windows.FontWeights]::Bold
                    Foreground      = $script:BrushConverter.ConvertFromString($p.AccentBlue)
                    Background      = [System.Windows.Media.Brushes]::Transparent
                    BorderThickness = New-Object System.Windows.Thickness(0)
                    Padding         = New-Object System.Windows.Thickness(4, 0, 4, 0)
                    Margin          = New-Object System.Windows.Thickness(4, 0, 0, 0)
                    Cursor          = [System.Windows.Input.Cursors]::Hand
                    Visibility      = if ($isDirty) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
                }

                $capturedChg = $chg
                $capturedBdr = $bdrNew
                $capturedRevert = $btnRevert
                $capturedTxt = $txtEditNew

                $txtEditNew.add_TextChanged({
                    $curText = $capturedTxt.Text
                    $origVal = if ($capturedChg.Contains('OriginalNewValue') -and $null -ne $capturedChg.OriginalNewValue) { $capturedChg.OriginalNewValue.ToString() } else { '' }
                    $capturedChg.NewValue = $curText
                    if ($curText -ne $origVal) {
                        $capturedChg.CustomEdited = $true
                        $capturedBdr.BorderBrush = $script:BrushConverter.ConvertFromString($p.AccentBlue)
                        $capturedBdr.BorderThickness = New-Object System.Windows.Thickness(2)
                        $capturedRevert.Visibility = [System.Windows.Visibility]::Visible
                    } else {
                        $capturedChg.CustomEdited = $false
                        $capturedBdr.BorderBrush = $script:BrushConverter.ConvertFromString($p.DiffCellNewBorder)
                        $capturedBdr.BorderThickness = New-Object System.Windows.Thickness(1)
                        $capturedRevert.Visibility = [System.Windows.Visibility]::Collapsed
                    }
                    & $script:UpdateStagingSummary
                })

                $btnRevert.add_Click({
                    $origVal = if ($capturedChg.Contains('OriginalNewValue') -and $null -ne $capturedChg.OriginalNewValue) { $capturedChg.OriginalNewValue.ToString() } else { '' }
                    $capturedTxt.Text = $origVal
                    $capturedChg.NewValue = $origVal
                    $capturedChg.CustomEdited = $false
                    $capturedBdr.BorderBrush = $script:BrushConverter.ConvertFromString($p.DiffCellNewBorder)
                    $capturedBdr.BorderThickness = New-Object System.Windows.Thickness(1)
                    $capturedRevert.Visibility = [System.Windows.Visibility]::Collapsed
                    & $script:UpdateStagingSummary
                })

                [System.Windows.Controls.Grid]::SetColumn($txtEditNew, 0)
                [System.Windows.Controls.Grid]::SetColumn($btnRevert, 1)
                [void]$gridNewCell.Children.Add($txtEditNew)
                [void]$gridNewCell.Children.Add($btnRevert)
                $bdrNew.Child = $gridNewCell

                [System.Windows.Controls.Grid]::SetRow($chk, $rIdx); [System.Windows.Controls.Grid]::SetColumn($chk, 0); [void]$grid.Children.Add($chk)
                [System.Windows.Controls.Grid]::SetRow($txtCol, $rIdx); [System.Windows.Controls.Grid]::SetColumn($txtCol, 1); [void]$grid.Children.Add($txtCol)
                [System.Windows.Controls.Grid]::SetRow($txtOld, $rIdx); [System.Windows.Controls.Grid]::SetColumn($txtOld, 2); [void]$grid.Children.Add($txtOld)
                [System.Windows.Controls.Grid]::SetRow($bdrNew, $rIdx); [System.Windows.Controls.Grid]::SetColumn($bdrNew, 3); [void]$grid.Children.Add($bdrNew)
                $rIdx++
            }
            [void]$panelDiffContainer.Children.Add($grid)

        } elseif ($status -eq 'Unchanged') {
            $baseRowStr = if ($rec.MatchedBaseRow) { (Get-UiString 'RowNumberFormat') -f $rec.MatchedBaseRow.RowNumber } else { '' }
            $txtDetailHeader.Text = "$((Get-UiString 'StatusBadgeUnchanged')): $($Item.Title) ($baseRowStr)"

            if ($rec.Changes -and $rec.Changes.Count -gt 0) {
                # Row Action Tools (Select/Deselect All in Row)
                $spRowTools = New-Object System.Windows.Controls.StackPanel -Property @{
                    Orientation = [System.Windows.Controls.Orientation]::Horizontal
                    Margin      = New-Object System.Windows.Thickness(0, 0, 0, 8)
                }
                $btnSelAllRow = New-Object System.Windows.Controls.Button -Property @{
                    Content         = Get-UiString 'BtnSelectAllRowFields' 'Zaznacz wszystkie w wierszu'
                    Background      = $script:BrushConverter.ConvertFromString($p.BgCardHover)
                    Foreground      = $script:BrushConverter.ConvertFromString($p.TextPrimary)
                    BorderBrush     = $script:BrushConverter.ConvertFromString($p.BorderCard)
                    BorderThickness = New-Object System.Windows.Thickness(1)
                    Padding         = New-Object System.Windows.Thickness(8, 3, 8, 3)
                    Margin          = New-Object System.Windows.Thickness(0, 0, 6, 0)
                    FontSize        = 11
                    Cursor          = [System.Windows.Input.Cursors]::Hand
                }
                $btnDeselAllRow = New-Object System.Windows.Controls.Button -Property @{
                    Content         = Get-UiString 'BtnDeselectAllRowFields' 'Odznacz wszystkie w wierszu'
                    Background      = $script:BrushConverter.ConvertFromString($p.BgCardHover)
                    Foreground      = $script:BrushConverter.ConvertFromString($p.TextPrimary)
                    BorderBrush     = $script:BrushConverter.ConvertFromString($p.BorderCard)
                    BorderThickness = New-Object System.Windows.Thickness(1)
                    Padding         = New-Object System.Windows.Thickness(8, 3, 8, 3)
                    FontSize        = 11
                    Cursor          = [System.Windows.Input.Cursors]::Hand
                }
                $allRowCheckboxes = [System.Collections.Generic.List[object]]::new()
                $btnSelAllRow.add_Click({
                    foreach ($c in $allRowCheckboxes) { $c.IsChecked = $true }
                    & $script:UpdateStagingSummary
                })
                $btnDeselAllRow.add_Click({
                    foreach ($c in $allRowCheckboxes) { $c.IsChecked = $false }
                    & $script:UpdateStagingSummary
                })
                [void]$spRowTools.Children.Add($btnSelAllRow)
                [void]$spRowTools.Children.Add($btnDeselAllRow)
                [void]$panelDiffContainer.Children.Add($spRowTools)

                # Diff table with selective update and inline editing
                $grid = New-Object System.Windows.Controls.Grid
                $grid.Margin = New-Object System.Windows.Thickness(0, 0, 0, 0)
                [void]$grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::Auto }))
                [void]$grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(180) }))
                [void]$grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) }))
                [void]$grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) }))

                [void]$grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto }))
                $hApp = New-Object System.Windows.Controls.TextBlock -Property @{ Text = Get-UiString 'ColApply'; Foreground = $script:BrushConverter.ConvertFromString($p.TextSecondary); FontWeight = [System.Windows.FontWeights]::Bold; Margin = New-Object System.Windows.Thickness(0, 0, 10, 8) }
                $hFld = New-Object System.Windows.Controls.TextBlock -Property @{ Text = Get-UiString 'ColField'; Foreground = $script:BrushConverter.ConvertFromString($p.TextSecondary); FontWeight = [System.Windows.FontWeights]::Bold; Margin = New-Object System.Windows.Thickness(0, 0, 10, 8) }
                $hCur = New-Object System.Windows.Controls.TextBlock -Property @{ Text = Get-UiString 'ColCurrentBase'; Foreground = $script:BrushConverter.ConvertFromString($p.TextSecondary); FontWeight = [System.Windows.FontWeights]::Bold; Margin = New-Object System.Windows.Thickness(0, 0, 10, 8) }
                $hInc = New-Object System.Windows.Controls.TextBlock -Property @{ Text = Get-UiString 'ColIncomingValue'; Foreground = $script:BrushConverter.ConvertFromString($p.TextSecondary); FontWeight = [System.Windows.FontWeights]::Bold; Margin = New-Object System.Windows.Thickness(0, 0, 0, 8) }

                [System.Windows.Controls.Grid]::SetRow($hApp, 0); [System.Windows.Controls.Grid]::SetColumn($hApp, 0); [void]$grid.Children.Add($hApp)
                [System.Windows.Controls.Grid]::SetRow($hFld, 0); [System.Windows.Controls.Grid]::SetColumn($hFld, 1); [void]$grid.Children.Add($hFld)
                [System.Windows.Controls.Grid]::SetRow($hCur, 0); [System.Windows.Controls.Grid]::SetColumn($hCur, 2); [void]$grid.Children.Add($hCur)
                [System.Windows.Controls.Grid]::SetRow($hInc, 0); [System.Windows.Controls.Grid]::SetColumn($hInc, 3); [void]$grid.Children.Add($hInc)

                $rIdx = 1
                foreach ($chg in $rec.Changes) {
                    [void]$grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto }))
                    $colKey = $chg.BaseColumn
                    $chk = New-Object System.Windows.Controls.CheckBox -Property @{
                        IsChecked = if ($Item.SelectedCells) { ($Item.SelectedCells[$colKey] -ne $false) } else { $true }
                        VerticalAlignment = [System.Windows.VerticalAlignment]::Center
                        Margin = New-Object System.Windows.Thickness(0, 4, 10, 4)
                    }
                    [void]$allRowCheckboxes.Add($chk)
                    $chk.add_Checked({
                        $Item.SelectedCells[$colKey] = $true
                        $chg.SelectedForUpdate = $true
                        & $script:UpdateStagingSummary
                    })
                    $chk.add_Unchecked({
                        $Item.SelectedCells[$colKey] = $false
                        $chg.SelectedForUpdate = $false
                        & $script:UpdateStagingSummary
                    })

                    $txtCol = New-Object System.Windows.Controls.TextBlock -Property @{
                        Text = $chg.BaseColumn
                        Foreground = $script:BrushConverter.ConvertFromString($p.TextPrimary)
                        FontWeight = [System.Windows.FontWeights]::SemiBold
                        VerticalAlignment = [System.Windows.VerticalAlignment]::Center
                        Margin = New-Object System.Windows.Thickness(0, 4, 10, 4)
                    }
                    $txtOld = New-Object System.Windows.Controls.TextBlock -Property @{
                        Text = $(if ($null -ne $chg.OldValue) { $chg.OldValue.ToString() } else { '' })
                        Foreground = $script:BrushConverter.ConvertFromString($p.DiffCellOldFg)
                        VerticalAlignment = [System.Windows.VerticalAlignment]::Center
                        Margin = New-Object System.Windows.Thickness(0, 4, 10, 4)
                    }

                    $isDirty = ($chg.Contains('CustomEdited') -and $chg.CustomEdited -eq $true)
                    $bdrBrushHex = if ($isDirty) { $p.AccentBlue } else { $p.DiffCellNewBorder }
                    $bdrThickVal = if ($isDirty) { 2 } else { 1 }
                    $bdrNew = New-Object System.Windows.Controls.Border -Property @{
                        Background      = $script:BrushConverter.ConvertFromString($p.DiffCellNewBg)
                        BorderBrush     = $script:BrushConverter.ConvertFromString($bdrBrushHex)
                        BorderThickness = New-Object System.Windows.Thickness($bdrThickVal)
                        CornerRadius    = New-Object System.Windows.CornerRadius(4)
                        Padding         = New-Object System.Windows.Thickness(4, 2, 4, 2)
                        Margin          = New-Object System.Windows.Thickness(0, 2, 0, 2)
                    }

                    $gridNewCell = New-Object System.Windows.Controls.Grid
                    [void]$gridNewCell.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) }))
                    [void]$gridNewCell.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::Auto }))

                    $txtEditNew = New-Object System.Windows.Controls.TextBox -Property @{
                        Text            = $(if ($null -ne $chg.NewValue) { $chg.NewValue.ToString() } else { '' })
                        Foreground      = $script:BrushConverter.ConvertFromString($p.DiffCellNewFg)
                        Background      = [System.Windows.Media.Brushes]::Transparent
                        BorderThickness = New-Object System.Windows.Thickness(0)
                        FontWeight      = [System.Windows.FontWeights]::Bold
                        VerticalAlignment = [System.Windows.VerticalAlignment]::Center
                        Padding         = New-Object System.Windows.Thickness(2, 1, 2, 1)
                    }

                    $btnRevert = New-Object System.Windows.Controls.Button -Property @{
                        Content         = '↺'
                        ToolTip         = (Get-UiString 'TooltipRevertOriginal' 'Przywróć wartość z pliku zmian')
                        FontSize        = 12
                        FontWeight      = [System.Windows.FontWeights]::Bold
                        Foreground      = $script:BrushConverter.ConvertFromString($p.AccentBlue)
                        Background      = [System.Windows.Media.Brushes]::Transparent
                        BorderThickness = New-Object System.Windows.Thickness(0)
                        Padding         = New-Object System.Windows.Thickness(4, 0, 4, 0)
                        Margin          = New-Object System.Windows.Thickness(4, 0, 0, 0)
                        Cursor          = [System.Windows.Input.Cursors]::Hand
                        Visibility      = if ($isDirty) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
                    }

                    $capturedChg = $chg
                    $capturedBdr = $bdrNew
                    $capturedRevert = $btnRevert
                    $capturedTxt = $txtEditNew

                    $txtEditNew.add_TextChanged({
                        $curText = $capturedTxt.Text
                        $origVal = if ($capturedChg.Contains('OriginalNewValue') -and $null -ne $capturedChg.OriginalNewValue) { $capturedChg.OriginalNewValue.ToString() } else { '' }
                        $capturedChg.NewValue = $curText
                        if ($curText -ne $origVal) {
                            $capturedChg.CustomEdited = $true
                            $capturedBdr.BorderBrush = $script:BrushConverter.ConvertFromString($p.AccentBlue)
                            $capturedBdr.BorderThickness = New-Object System.Windows.Thickness(2)
                            $capturedRevert.Visibility = [System.Windows.Visibility]::Visible
                        } else {
                            $capturedChg.CustomEdited = $false
                            $capturedBdr.BorderBrush = $script:BrushConverter.ConvertFromString($p.DiffCellNewBorder)
                            $capturedBdr.BorderThickness = New-Object System.Windows.Thickness(1)
                            $capturedRevert.Visibility = [System.Windows.Visibility]::Collapsed
                        }
                        & $script:UpdateStagingSummary
                    })

                    $btnRevert.add_Click({
                        $origVal = if ($capturedChg.Contains('OriginalNewValue') -and $null -ne $capturedChg.OriginalNewValue) { $capturedChg.OriginalNewValue.ToString() } else { '' }
                        $capturedTxt.Text = $origVal
                        $capturedChg.NewValue = $origVal
                        $capturedChg.CustomEdited = $false
                        $capturedBdr.BorderBrush = $script:BrushConverter.ConvertFromString($p.DiffCellNewBorder)
                        $capturedBdr.BorderThickness = New-Object System.Windows.Thickness(1)
                        $capturedRevert.Visibility = [System.Windows.Visibility]::Collapsed
                        & $script:UpdateStagingSummary
                    })

                    [System.Windows.Controls.Grid]::SetColumn($txtEditNew, 0)
                    [System.Windows.Controls.Grid]::SetColumn($btnRevert, 1)
                    [void]$gridNewCell.Children.Add($txtEditNew)
                    [void]$gridNewCell.Children.Add($btnRevert)
                    $bdrNew.Child = $gridNewCell

                    [System.Windows.Controls.Grid]::SetRow($chk, $rIdx); [System.Windows.Controls.Grid]::SetColumn($chk, 0); [void]$grid.Children.Add($chk)
                    [System.Windows.Controls.Grid]::SetRow($txtCol, $rIdx); [System.Windows.Controls.Grid]::SetColumn($txtCol, 1); [void]$grid.Children.Add($txtCol)
                    [System.Windows.Controls.Grid]::SetRow($txtOld, $rIdx); [System.Windows.Controls.Grid]::SetColumn($txtOld, 2); [void]$grid.Children.Add($txtOld)
                    [System.Windows.Controls.Grid]::SetRow($bdrNew, $rIdx); [System.Windows.Controls.Grid]::SetColumn($bdrNew, 3); [void]$grid.Children.Add($bdrNew)
                    $rIdx++
                }
                [void]$panelDiffContainer.Children.Add($grid)
            } else {
                $infoTxt = New-Object System.Windows.Controls.TextBlock -Property @{
                    Text = Get-UiString 'UnchangedRowNotice'
                    Foreground = $script:BrushConverter.ConvertFromString($p.TextSecondary)
                    FontSize = 13
                    Margin = New-Object System.Windows.Thickness(0, 12, 0, 12)
                    TextWrapping = [System.Windows.TextWrapping]::Wrap
                }
                [void]$panelDiffContainer.Children.Add($infoTxt)
            }

        } elseif ($status -eq 'Ambiguous') {
            $txtDetailHeader.Text = "$((Get-UiString 'StatusBadgeAmbiguous')): $($Item.Title)"

            $alertBdr = New-Object System.Windows.Controls.Border -Property @{
                Background = $script:BrushConverter.ConvertFromString($p.AlertAmbBg)
                BorderBrush = $script:BrushConverter.ConvertFromString($p.AlertAmbBorder)
                BorderThickness = New-Object System.Windows.Thickness(1)
                CornerRadius = New-Object System.Windows.CornerRadius(4)
                Padding = New-Object System.Windows.Thickness(12)
                Margin = New-Object System.Windows.Thickness(0, 0, 0, 12)
            }
            $alertSp = New-Object System.Windows.Controls.StackPanel
            $alertHdr = New-Object System.Windows.Controls.TextBlock -Property @{
                Text = Get-UiString 'AmbiguousHeader'
                Foreground = $script:BrushConverter.ConvertFromString($p.AlertAmbHdrFg)
                FontWeight = [System.Windows.FontWeights]::Bold
                FontSize = 14
                Margin = New-Object System.Windows.Thickness(0, 0, 0, 4)
            }
            $alertSub = New-Object System.Windows.Controls.TextBlock -Property @{
                Text = Get-UiString 'AmbiguousSelectPrompt'
                Foreground = $script:BrushConverter.ConvertFromString($p.AlertAmbSubFg)
                FontSize = 12
            }
            [void]$alertSp.Children.Add($alertHdr)
            [void]$alertSp.Children.Add($alertSub)
            $alertBdr.Child = $alertSp
            [void]$panelDiffContainer.Children.Add($alertBdr)

            # Candidates List
            $cIdx = 1
            $candidateRows = if ($rec.CandidateBaseRows) { $rec.CandidateBaseRows } elseif ($rec.Candidates) { $rec.Candidates } else { @() }
            foreach ($cand in $candidateRows) {
                $candBdr = New-Object System.Windows.Controls.Border -Property @{
                    Background = $script:BrushConverter.ConvertFromString($p.BgCardHover)
                    BorderBrush = $script:BrushConverter.ConvertFromString($p.BorderCard)
                    BorderThickness = New-Object System.Windows.Thickness(1)
                    CornerRadius = New-Object System.Windows.CornerRadius(4)
                    Padding = New-Object System.Windows.Thickness(10)
                    Margin = New-Object System.Windows.Thickness(0, 0, 0, 8)
                }
                $candSp = New-Object System.Windows.Controls.StackPanel -Property @{ Orientation = [System.Windows.Controls.Orientation]::Horizontal }
                $candBtn = New-Object System.Windows.Controls.Button -Property @{
                    Content = (Get-UiString 'BtnSelectCandidate') -f $cIdx, ((Get-UiString 'RowNumberFormat') -f $cand.RowNumber)
                    Background = $script:BrushConverter.ConvertFromString($p.AccentBlue)
                    Foreground = $script:BrushConverter.ConvertFromString('#FFFFFF')
                    FontWeight = [System.Windows.FontWeights]::Bold
                    Padding = New-Object System.Windows.Thickness(12, 6, 12, 6)
                    Margin = New-Object System.Windows.Thickness(0, 0, 12, 0)
                }

                $capturedCand = $cand
                $candBtn.add_Click({
                    $rec.MatchedBaseRow = $capturedCand
                    $rec.Status = 'Changed'
                    $Item.StatusText = Get-UiString 'StatusBadgeChanged'
                    $bColors = Get-StatusBadgeColors 'Changed' $script:CurrentTheme
                    $Item.StatusBg = $bColors.Bg
                    $Item.StatusFg = $bColors.Fg
                    # Recompute diffs with exact ColIndex and CellRef coordinates
                    $newDiffs = [System.Collections.Generic.List[object]]::new()
                    foreach ($k in $rec.ProjectedRow.Keys) {
                        $oldVal = if ($capturedCand.Values.ContainsKey($k)) { $capturedCand.Values[$k] } else { '' }
                        $newVal = $rec.ProjectedRow[$k]
                        if (-not [FastDiffHelper]::AreEqual($oldVal, $newVal, $true, $true, $true, $true)) {
                            $colIdx = if ($script:BaseHeaders) { [System.Array]::IndexOf($script:BaseHeaders, $k) } else { -1 }
                            $colLetter = if ($colIdx -ge 0) { [FastExcelHelper]::ColIndexToName($colIdx) } else { '' }
                            $cellRef = if ($colLetter) { "$colLetter$($capturedCand.RowNumber)" } else { '' }
                            $newDiffs.Add([PSCustomObject]@{
                                BaseColumn        = $k
                                ColIndex          = $colIdx
                                CellRef           = $cellRef
                                OldValue          = $oldVal
                                NewValue          = $newVal
                                SelectedForUpdate = $true
                            })
                            if (-not $Item.SelectedCells) { $Item.SelectedCells = @{} }; $Item.SelectedCells[$k] = $true
                        }
                    }
                    $rec.Changes = $newDiffs
                    & $RenderDetailPane $Item
                    & $UpdateCounters
                })

                $candSummaryValues = ($cand.Values.GetEnumerator() | Select-Object -First 4 | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ', '
                $summaryTxt = New-Object System.Windows.Controls.TextBlock -Property @{
                    Text = (Get-UiString 'CandidateRowSummary') -f $cand.RowNumber, $candSummaryValues
                    Foreground = $script:BrushConverter.ConvertFromString($p.TextPrimary)
                    VerticalAlignment = [System.Windows.VerticalAlignment]::Center
                }
                [void]$candSp.Children.Add($candBtn)
                [void]$candSp.Children.Add($summaryTxt)
                $candBdr.Child = $candSp
                [void]$panelDiffContainer.Children.Add($candBdr)
                $cIdx++
            }

            # Treat as New Row button
            $btnTreatNew = New-Object System.Windows.Controls.Button -Property @{
                Content = Get-UiString 'BtnTreatAsNew'
                Background = $script:BrushConverter.ConvertFromString($p.AccentGreen)
                Foreground = $script:BrushConverter.ConvertFromString('#FFFFFF')
                FontWeight = [System.Windows.FontWeights]::Bold
                Padding = New-Object System.Windows.Thickness(16, 8, 16, 8)
                HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
                Margin = New-Object System.Windows.Thickness(0, 8, 0, 0)
            }
            $btnTreatNew.add_Click({
                $rec.MatchedBaseRow = $null
                $rec.Changes = @()
                $rec.Status = 'New'
                $Item.StatusText = Get-UiString 'StatusBadgeNew'
                $bColors = Get-StatusBadgeColors 'New' $script:CurrentTheme
                $Item.StatusBg = $bColors.Bg
                $Item.StatusFg = $bColors.Fg
                & $RenderDetailPane $Item
                & $UpdateCounters
            })
            [void]$panelDiffContainer.Children.Add($btnTreatNew)
        }

        # Always update full base row preview in bottom card unless skipped
        if ($script:RenderFullBaseRow -and -not $SkipBaseRowRefresh) {
            & $script:RenderFullBaseRow $Item
        }
    }

    # Render Full Base Row DataGrid
    $script:RenderFullBaseRow = $RenderFullBaseRow = {
        param($Item)
        if (-not $baseRowBorder) {
            $baseRowBorder = if ($window) { $window.FindName('baseRowBorder') } else { $script:ActiveWindow.FindName('baseRowBorder') }
        }
        if (-not $dgBaseFullRow) {
            $dgBaseFullRow = if ($window) { $window.FindName('dgBaseFullRow') } else { $script:ActiveWindow.FindName('dgBaseFullRow') }
        }
        if (-not $txtBaseRowHeader) {
            $txtBaseRowHeader = if ($window) { $window.FindName('txtBaseRowHeader') } else { $script:ActiveWindow.FindName('txtBaseRowHeader') }
        }
        if (-not $Item -or -not $script:BaseHeaders -or $script:BaseHeaders.Count -eq 0) {
            if ($baseRowBorder) { $baseRowBorder.Visibility = [System.Windows.Visibility]::Collapsed }
            return
        }

        if ($Item -and -not $Item.SelectedCells) { $Item.SelectedCells = @{} }
        $rec = $Item.Record
        $hasBaseRow = ($null -ne $rec.MatchedBaseRow)
        $isNew = ($rec.Status -eq 'New')

        if (-not $hasBaseRow -and -not $isNew) {
            if ($baseRowBorder) { $baseRowBorder.Visibility = [System.Windows.Visibility]::Collapsed }
            return
        }

        if ($baseRowBorder) { $baseRowBorder.Visibility = [System.Windows.Visibility]::Visible }

        if ($hasBaseRow) {
            $txtBaseRowHeader.Text = (Get-UiString 'BaseRowHeader') -f $rec.MatchedBaseRow.RowNumber
        } else {
            $txtBaseRowHeader.Text = Get-UiString 'BaseRowHeaderNew'
        }

        $dt = New-Object System.Data.DataTable
        $dgBaseFullRow.Columns.Clear()

        foreach ($h in $script:BaseHeaders) {
            $colKey = $h
            if (-not $dt.Columns.Contains($colKey)) {
                [void]$dt.Columns.Add($colKey, [string])
            }
            $col = New-Object System.Windows.Controls.DataGridTextColumn -Property @{
                Header     = $h
                Binding    = New-Object System.Windows.Data.Binding("[$colKey]")
                IsReadOnly = $false
            }
            $dgBaseFullRow.Columns.Add($col)
        }

        $dr = $dt.NewRow()
        foreach ($h in $script:BaseHeaders) {
            if ($dt.Columns.Contains($h)) {
                $val = ''
                if ($hasBaseRow -and $rec.MatchedBaseRow.Values) {
                    if ($rec.MatchedBaseRow.Values.ContainsKey($h)) {
                        $val = $rec.MatchedBaseRow.Values[$h]
                    }
                } elseif ($isNew -and $rec.ProjectedRow) {
                    if ($rec.ProjectedRow.ContainsKey($h)) {
                        $val = $rec.ProjectedRow[$h]
                    }
                }
                if ($rec.Changes) {
                    $cMatch = $rec.Changes | Where-Object { $_.BaseColumn -eq $h } | Select-Object -First 1
                    if ($cMatch -and $null -ne $cMatch.NewValue) {
                        $val = $cMatch.NewValue
                    }
                }
                $dr[$h] = if ($null -ne $val) { $val.ToString() } else { '' }
            }
        }
        $dt.Rows.Add($dr)
        $dgBaseFullRow.ItemsSource = $dt.DefaultView
    }

    $dgBaseFullRow.add_CellEditEnding({
        param($sender, $e)
        if ($e.EditAction -eq [System.Windows.Controls.DataGridEditAction]::Cancel) { return }
        $sel = $lbReviewItems.SelectedItem
        if (-not $sel) { return }

        $colName = if ($e.Column -and $e.Column.Header) { $e.Column.Header.ToString() } else { '' }
        if ([string]::IsNullOrEmpty($colName)) { return }

        $tb = $e.EditingElement -as [System.Windows.Controls.TextBox]
        $newVal = if ($tb) { $tb.Text } else { '' }

        $rec = $sel.Record
        $oldVal = ''
        if ($rec.MatchedBaseRow -and $rec.MatchedBaseRow.Values) {
            $oldVal = if ($rec.MatchedBaseRow.Values.ContainsKey($colName)) { $rec.MatchedBaseRow.Values[$colName] } else { '' }
            $rec.MatchedBaseRow.Values[$colName] = $newVal
        } elseif ($rec.ProjectedRow) {
            $oldVal = if ($rec.ProjectedRow.ContainsKey($colName)) { $rec.ProjectedRow[$colName] } else { '' }
            $rec.ProjectedRow[$colName] = $newVal
        }

        if ($null -eq $rec.Changes -or $rec.Changes -isnot [System.Collections.Generic.List[object]]) {
            $rec.Changes = [System.Collections.Generic.List[object]]::new([object[]]@($rec.Changes))
        }
        $existing = $rec.Changes | Where-Object { $_.BaseColumn -eq $colName } | Select-Object -First 1
        if ($existing) {
            $existing.NewValue = $newVal
            $existing.SelectedForUpdate = $true
        } else {
            $rec.Changes.Add([PSCustomObject]@{
                BaseColumn        = $colName
                OldValue          = $oldVal
                NewValue          = $newVal
                SelectedForUpdate = $true
            })
        }
        if (-not $sel.SelectedCells) { $sel.SelectedCells = @{} }; $sel.SelectedCells[$colName] = $true

        if ($rec.Status -eq 'Unchanged') {
            $rec.Status = 'Changed'
            $sel.StatusText = Get-UiString 'StatusBadgeChanged'
            $bColors = Get-StatusBadgeColors 'Changed' $script:CurrentTheme
            $sel.StatusBg = $bColors.Bg
            $sel.StatusFg = $bColors.Fg
            & $UpdateCounters
        }

        $sel.Subtitle = (Get-UiString 'SubtitleFieldChanges') -f $rec.Changes.Count

        # Defer all UI updates to after CellEditEnding completes so that:
        # 1. The edited cell value is fully committed before re-render
        # 2. Items.Refresh() does not fire SelectionChanged mid-edit (which would wipe base row)
        $targetItem = $sel
        $capturedAllItems = $script:AllReviewItems
        $window.Dispatcher.InvokeAsync([System.Action]{
            # Update subtitle (Changes count may have grown)
            $itemToRender = if ($lbReviewItems -and $lbReviewItems.SelectedItem) { $lbReviewItems.SelectedItem } else { $targetItem }

            # Refresh list display without losing selection
            $curIdx = if ($lbReviewItems) { $lbReviewItems.SelectedIndex } else { -1 }
            if ($lbReviewItems) { $lbReviewItems.Items.Refresh() }
            if ($curIdx -ge 0 -and $lbReviewItems) {
                # Suppress SelectionChanged re-render while we restore the same selection index
                $script:_suppressSelectionRender = $true
                $lbReviewItems.SelectedIndex = $curIdx
                $script:_suppressSelectionRender = $false
            }

            # Update Apply button state
            if ($btnApplyAccepted) {
                $btnApplyAccepted.IsEnabled = ($capturedAllItems | Where-Object { $_.Decision -eq 'Accepted' }).Count -gt 0
            }

            # Re-render detail pane with updated changes, keep base row intact
            if ($itemToRender -and $script:RenderDetailPane) {
                & $script:RenderDetailPane $itemToRender -SkipBaseRowRefresh
            }
        }.GetNewClosure())
    })

    # Selection changed on Review list
    $lbReviewItems.add_SelectionChanged({
        if ($script:_suppressSelectionRender) { return }
        if ($lbReviewItems.SelectedItem) {
            & $RenderDetailPane $lbReviewItems.SelectedItem
            $lbReviewItems.ScrollIntoView($lbReviewItems.SelectedItem)
        }
    })

    # Re-click on already selected review item also refreshes detail pane
    $lbReviewItems.add_PreviewMouseLeftButtonUp({
        if ($lbReviewItems.SelectedItem) {
            & $RenderDetailPane $lbReviewItems.SelectedItem
        }
    })

    # Filter & Search logic
    $ApplyFilterAndSearch = {
        if (-not $script:AllReviewItems) { return }
        $filterTag = if ($cmbFilterStatus.SelectedItem -is [System.Windows.Controls.ComboBoxItem]) {
            $cmbFilterStatus.SelectedItem.Tag
        } else { 'All' }

        $searchQuery = if ($txtSearchReview.Text) { $txtSearchReview.Text.Trim().ToLowerInvariant() } else { '' }

        $filtered = $script:AllReviewItems | Where-Object {
            $item = $_
            $matchStatus = switch ($filterTag) {
                'All'       { $true }
                'New'       { $item.Record.Status -eq 'New' }
                'Changed'   { $item.Record.Status -eq 'Changed' }
                'Removed'   { $item.Record.Status -eq 'Removed' }
                'Ambiguous' { $item.Record.Status -eq 'Ambiguous' }
                'Unchanged' { $item.Record.Status -eq 'Unchanged' }
                'Accepted'  { $item.Decision -eq 'Accepted' }
                'Skipped'   { $item.Decision -eq 'Skipped' }
                'Rejected'  { $item.Decision -eq 'Rejected' }
                default     { $true }
            }
            if (-not $matchStatus) { return $false }
            if ([string]::IsNullOrWhiteSpace($searchQuery)) { return $true }

            $titleMatch = if ($item.Title) { $item.Title.ToLowerInvariant().Contains($searchQuery) } else { $false }
            $subMatch   = if ($item.Subtitle) { $item.Subtitle.ToLowerInvariant().Contains($searchQuery) } else { $false }
            return ($titleMatch -or $subMatch)
        }

        $lbReviewItems.ItemsSource = @($filtered)
        if ($lbReviewItems.Items.Count -gt 0) {
            $lbReviewItems.SelectedIndex = 0
        }
    }
    $script:ApplyFilterAndSearch = $ApplyFilterAndSearch

    # Populate Review Items Controller
    $script:PopulateReviewItems = {
        if (-not $script:ComparisonResult) { return }

        $incKeys = [string[]]@($script:MappingRules | Where-Object { $_.IsJoinKey } | ForEach-Object { $_.IncomingColumn })
        $baseKeys = [string[]]@($script:MappingRules | Where-Object { $_.IsJoinKey } | ForEach-Object { $_.BaseColumn })
        $showUnchanged = if ($chkShowUnchanged) { ($chkShowUnchanged.IsChecked -eq $true) } else { ($script:AppConfig.ShowUnchangedRows -eq $true -or $script:AppConfig.AutoSkipUnchanged -eq $false) }

        # Preserve existing decisions and cell selections across toggles
        $existingDecisions = @{}
        $existingCells     = @{}
        if ($script:AllReviewItems) {
            foreach ($it in $script:AllReviewItems) {
                if ($it.Record) {
                    $existingDecisions[$it.Record] = $it.Decision
                    $existingCells[$it.Record]     = $it.SelectedCells
                }
            }
        }

        $script:AllReviewItems.Clear()
        $idx = 1
        foreach ($r in $script:ComparisonResult) {
            if (-not $showUnchanged -and $r.Status -eq 'Unchanged') {
                continue
            }

            $title = if ($r.IncomingRow -and $incKeys -and $incKeys.Length -gt 0 -and $r.IncomingRow.Values -and $r.IncomingRow.Values.ContainsKey($incKeys[0])) {
                $r.IncomingRow.Values[$incKeys[0]].ToString()
            } elseif ($r.IncomingRow -and $incKeys.Length -gt 0 -and $r.IncomingRow.PSObject.Properties[$incKeys[0]]) {
                $r.IncomingRow.$($incKeys[0]).ToString()
            } elseif ($r.MatchedBaseRow -and $baseKeys.Length -gt 0 -and $r.MatchedBaseRow.Values -and $r.MatchedBaseRow.Values.ContainsKey($baseKeys[0])) {
                $r.MatchedBaseRow.Values[$baseKeys[0]].ToString()
            } elseif ($r.MatchedBaseRow -and $baseKeys.Length -gt 0 -and $r.MatchedBaseRow.PSObject.Properties[$baseKeys[0]]) {
                $r.MatchedBaseRow.$($baseKeys[0]).ToString()
            } else { (Get-UiString 'RowNumberFormat') -f $idx }

            $subtitle = switch ($r.Status) {
                'New'       { (Get-UiString 'SubtitleNew') -f $r.ProjectedRow.Count }
                'Changed'   { (Get-UiString 'SubtitleChanged') -f $r.Changes.Count }
                'Removed'   {
                    $col = if ($r.Changes -and $r.Changes.Count -gt 0) { $r.Changes[0].BaseColumn } else { '' }
                    $val = if ($r.Changes -and $r.Changes.Count -gt 0) { $r.Changes[0].NewValue } else { '' }
                    (Get-UiString 'SubtitleRemoved') -f $col, $val
                }
                'Ambiguous' {
                    $cCnt = if ($r.CandidateBaseRows) { $r.CandidateBaseRows.Count } elseif ($r.Candidates) { $r.Candidates.Count } else { 0 }
                    (Get-UiString 'SubtitleAmbiguous') -f $cCnt
                }
                'Unchanged' { Get-UiString 'SubtitleUnchanged' }
                default     { $r.Status }
            }

            $statusText = switch ($r.Status) {
                'New'       { Get-UiString 'StatusBadgeNew' }
                'Changed'   { Get-UiString 'StatusBadgeChanged' }
                'Removed'   { Get-UiString 'StatusBadgeRemoved' }
                'Ambiguous' { Get-UiString 'StatusBadgeAmbiguous' }
                'Unchanged' { Get-UiString 'StatusBadgeUnchanged' }
                default     { $r.Status }
            }

            $badge = Get-StatusBadgeColors -Status $r.Status -ThemeName $script:CurrentTheme
            $statusBg = $badge.Bg
            $statusFg = $badge.Fg

            $initialDecision = switch ($r.Status) {
                'New'       { 'Accepted' }
                'Changed'   { 'Accepted' }
                'Removed'   { 'Accepted' }
                'Unchanged' { 'Skipped' }
                default     { 'Pending' }
            }
            if ($existingDecisions.ContainsKey($r)) {
                $initialDecision = $existingDecisions[$r]
            }

            $cellMap = @{}
            if ($existingCells.ContainsKey($r)) {
                $cellMap = $existingCells[$r]
            } elseif ($r.Changes) {
                foreach ($c in $r.Changes) { $cellMap[$c.BaseColumn] = $true }
            }

            $script:AllReviewItems.Add([PSCustomObject]@{
                Record        = $r
                IndexStr      = "#$idx"
                Title         = $title
                Subtitle      = $subtitle
                StatusText    = $statusText
                StatusBg      = $statusBg
                StatusFg      = $statusFg
                Decision      = $initialDecision
                SelectedCells = $cellMap
            })
            $idx++
        }

        & $script:ApplyFilterAndSearch
        & $script:UpdateCounters
        if ($btnApplyAccepted) {
            $btnApplyAccepted.IsEnabled = ($script:AllReviewItems | Where-Object { $_.Decision -eq 'Accepted' }).Count -gt 0
        }

        # Populate Batch Column Toggles
        if ($wrapBatchColToggles -and $cardBatchToggles) {
            $wrapBatchColToggles.Children.Clear()
            $uniqueCols = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($it in $script:AllReviewItems) {
                if ($it.Record -and $it.Record.Changes) {
                    foreach ($c in $it.Record.Changes) {
                        if ($c.BaseColumn) { [void]$uniqueCols.Add($c.BaseColumn) }
                    }
                }
            }

            if ($uniqueCols.Count -gt 0) {
                $cardBatchToggles.Visibility = [System.Windows.Visibility]::Visible
                $p = Get-UpdaterThemePalette $script:CurrentTheme
                foreach ($colName in ($uniqueCols | Sort-Object)) {
                    $chkCol = New-Object System.Windows.Controls.CheckBox -Property @{
                        Content         = $colName
                        IsChecked       = $true
                        Foreground      = $script:BrushConverter.ConvertFromString($p.TextPrimary)
                        FontWeight      = [System.Windows.FontWeights]::SemiBold
                        FontSize        = 11
                        Margin          = New-Object System.Windows.Thickness(0, 0, 12, 4)
                        Cursor          = [System.Windows.Input.Cursors]::Hand
                    }
                    $capturedCol = $colName
                    $chkCol.add_Checked({
                        $colToToggle = $capturedCol
                        foreach ($item in $script:AllReviewItems) {
                            if ($item.SelectedCells) { $item.SelectedCells[$colToToggle] = $true }
                            if ($item.Record -and $item.Record.Changes) {
                                foreach ($c in $item.Record.Changes) {
                                    if ($c.BaseColumn -eq $colToToggle) { $c.SelectedForUpdate = $true }
                                }
                            }
                        }
                        if ($lbReviewItems.SelectedItem) { & $RenderDetailPane $lbReviewItems.SelectedItem }
                        & $script:UpdateStagingSummary
                    })
                    $chkCol.add_Unchecked({
                        $colToToggle = $capturedCol
                        foreach ($item in $script:AllReviewItems) {
                            if ($item.SelectedCells) { $item.SelectedCells[$colToToggle] = $false }
                            if ($item.Record -and $item.Record.Changes) {
                                foreach ($c in $item.Record.Changes) {
                                    if ($c.BaseColumn -eq $colToToggle) { $c.SelectedForUpdate = $false }
                                }
                            }
                        }
                        if ($lbReviewItems.SelectedItem) { & $RenderDetailPane $lbReviewItems.SelectedItem }
                        & $script:UpdateStagingSummary
                    })
                    [void]$wrapBatchColToggles.Children.Add($chkCol)
                }
            } else {
                $cardBatchToggles.Visibility = [System.Windows.Visibility]::Collapsed
            }
        }
    }

    # KPI Pill active selection updater
    $UpdateKpiPillSelection = {
        $tag = if ($cmbFilterStatus.SelectedItem -is [System.Windows.Controls.ComboBoxItem]) {
            $cmbFilterStatus.SelectedItem.Tag
        } elseif ($cmbFilterStatus.SelectedItem) {
            $cmbFilterStatus.SelectedItem.ToString()
        } else { 'All' }

        $pills = @(
            @{ Pill = $kpiPillAll;       Tag = 'All' },
            @{ Pill = $kpiPillNew;       Tag = 'New' },
            @{ Pill = $kpiPillChanged;   Tag = 'Changed' },
            @{ Pill = $kpiPillAmbiguous; Tag = 'Ambiguous' },
            @{ Pill = $kpiPillAccepted;  Tag = 'Accepted' },
            @{ Pill = $kpiPillSkipped;   Tag = 'Skipped' }
        )

        foreach ($p in $pills) {
            if ($p.Pill) {
                if ($p.Tag -eq $tag) {
                    $p.Pill.BorderBrush = $window.Resources['AccentBlue']
                    $p.Pill.BorderThickness = [System.Windows.Thickness]::new(2)
                } else {
                    $p.Pill.BorderBrush = $window.Resources['BorderCard']
                    $p.Pill.BorderThickness = [System.Windows.Thickness]::new(1)
                }
            }
        }
    }
    $script:UpdateKpiPillSelection = $UpdateKpiPillSelection

    # Click handlers for KPI filter pills
    $SelectFilterByTag = {
        param([string]$TargetTag)
        for ($i = 0; $i -lt $cmbFilterStatus.Items.Count; $i++) {
            $it = $cmbFilterStatus.Items[$i]
            $itemTag = if ($it -is [System.Windows.Controls.ComboBoxItem]) { $it.Tag } else { $it.ToString() }
            if ($itemTag -eq $TargetTag) {
                $cmbFilterStatus.SelectedIndex = $i
                break
            }
        }
    }

    if ($kpiPillAll)       { $kpiPillAll.add_MouseDown({ & $SelectFilterByTag 'All' }) }
    if ($kpiPillNew)       { $kpiPillNew.add_MouseDown({ & $SelectFilterByTag 'New' }) }
    if ($kpiPillChanged)   { $kpiPillChanged.add_MouseDown({ & $SelectFilterByTag 'Changed' }) }
    if ($kpiPillAmbiguous) { $kpiPillAmbiguous.add_MouseDown({ & $SelectFilterByTag 'Ambiguous' }) }
    if ($kpiPillAccepted)  { $kpiPillAccepted.add_MouseDown({ & $SelectFilterByTag 'Accepted' }) }
    if ($kpiPillSkipped)   { $kpiPillSkipped.add_MouseDown({ & $SelectFilterByTag 'Skipped' }) }

    $cmbFilterStatus.add_SelectionChanged({
        $selectedTag = if ($cmbFilterStatus.SelectedItem -is [System.Windows.Controls.ComboBoxItem]) {
            $cmbFilterStatus.SelectedItem.Tag
        } elseif ($cmbFilterStatus.SelectedItem) {
            $cmbFilterStatus.SelectedItem.ToString()
        } else { '' }

        if ($selectedTag -eq 'Unchanged' -and $chkShowUnchanged -and -not $chkShowUnchanged.IsChecked) {
            $chkShowUnchanged.IsChecked = $true
            $script:AppConfig.ShowUnchangedRows = $true
            $script:AppConfig.AutoSkipUnchanged = $false
            if ($script:PopulateReviewItems) { & $script:PopulateReviewItems }
        }

        if ($script:UpdateKpiPillSelection) { & $script:UpdateKpiPillSelection }
        if ($script:ApplyFilterAndSearch)   { & $script:ApplyFilterAndSearch }
    })

    if ($chkShowUnchanged) {
        $chkShowUnchanged.add_Click({
            $script:AppConfig.ShowUnchangedRows = ($chkShowUnchanged.IsChecked -eq $true)
            $script:AppConfig.AutoSkipUnchanged = ($chkShowUnchanged.IsChecked -ne $true)
            if ($script:PopulateReviewItems) { & $script:PopulateReviewItems }
        })
    }

    $txtSearchReview.add_TextChanged({
        if ($btnSearchClear) {
            $btnSearchClear.Visibility = if ([string]::IsNullOrEmpty($txtSearchReview.Text)) {
                [System.Windows.Visibility]::Collapsed
            } else {
                [System.Windows.Visibility]::Visible
            }
        }
        if ($script:ApplyFilterAndSearch) { & $script:ApplyFilterAndSearch }
    })

    if ($btnSearchClear) {
        $btnSearchClear.add_Click({
            $txtSearchReview.Text = ''
            $txtSearchReview.Focus()
        })
    }

    # Action Buttons (Back, Accept, Reject, Skip, Edit)
    $btnBackRow.add_Click({
        if ($lbReviewItems.SelectedIndex -gt 0) {
            $lbReviewItems.SelectedIndex--
            if ($lbReviewItems.SelectedItem) { $lbReviewItems.ScrollIntoView($lbReviewItems.SelectedItem) }
        }
    })

    $btnAcceptRow.add_Click({
        $sel = $lbReviewItems.SelectedItem
        if (-not $sel) { return }
        if ($sel.Record.Status -eq 'Ambiguous') {
            [System.Windows.Forms.MessageBox]::Show((Get-UiString 'AmbiguousSelectPrompt'), (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        # Push undo snapshot before changing decision
        $script:UndoStack.Push([PSCustomObject]@{ Item = $sel; PrevDecision = $sel.Decision; PrevStatusText = $sel.StatusText; PrevStatusBg = $sel.StatusBg; PrevStatusFg = $sel.StatusFg })
        $badge = Get-StatusBadgeColors -Status 'Accepted' -ThemeName $script:CurrentTheme
        $sel.Decision = 'Accepted'
        $sel.StatusText = Get-UiString 'StatusBadgeAccepted'
        $sel.StatusBg = $badge.Bg
        $sel.StatusFg = $badge.Fg

        if ($lbReviewItems.SelectedIndex -lt $lbReviewItems.Items.Count - 1) {
            $lbReviewItems.SelectedIndex++
        }
        & $UpdateCounters
        $btnApplyAccepted.IsEnabled = ($script:AllReviewItems | Where-Object { $_.Decision -eq 'Accepted' }).Count -gt 0
    })

    $btnRejectRow.add_Click({
        $sel = $lbReviewItems.SelectedItem
        if (-not $sel) { return }
        $script:UndoStack.Push([PSCustomObject]@{ Item = $sel; PrevDecision = $sel.Decision; PrevStatusText = $sel.StatusText; PrevStatusBg = $sel.StatusBg; PrevStatusFg = $sel.StatusFg })
        $badge = Get-StatusBadgeColors -Status 'Rejected' -ThemeName $script:CurrentTheme
        $sel.Decision = 'Rejected'
        $sel.StatusText = Get-UiString 'StatusBadgeRejected'
        $sel.StatusBg = $badge.Bg
        $sel.StatusFg = $badge.Fg

        if ($lbReviewItems.SelectedIndex -lt $lbReviewItems.Items.Count - 1) {
            $lbReviewItems.SelectedIndex++
        }
        & $UpdateCounters
        $btnApplyAccepted.IsEnabled = ($script:AllReviewItems | Where-Object { $_.Decision -eq 'Accepted' }).Count -gt 0
    })

    $btnSkipRow.add_Click({
        $sel = $lbReviewItems.SelectedItem
        if (-not $sel) { return }
        $script:UndoStack.Push([PSCustomObject]@{ Item = $sel; PrevDecision = $sel.Decision; PrevStatusText = $sel.StatusText; PrevStatusBg = $sel.StatusBg; PrevStatusFg = $sel.StatusFg })
        $badge = Get-StatusBadgeColors -Status 'Skipped' -ThemeName $script:CurrentTheme
        $sel.Decision = 'Skipped'
        $sel.StatusText = Get-UiString 'StatusBadgeSkipped'
        $sel.StatusBg = $badge.Bg
        $sel.StatusFg = $badge.Fg

        if ($lbReviewItems.SelectedIndex -lt $lbReviewItems.Items.Count - 1) {
            $lbReviewItems.SelectedIndex++
        }
        & $UpdateCounters
        $btnApplyAccepted.IsEnabled = ($script:AllReviewItems | Where-Object { $_.Decision -eq 'Accepted' }).Count -gt 0
    })

    # Open Modify Field Dialog (allows editing any field from base file, even if not in change file)
    $OpenModifyFieldDialog = {
        $sel = $lbReviewItems.SelectedItem
        if (-not $sel) {
            [System.Windows.Forms.MessageBox]::Show((Get-UiString 'NoRowSelectedForBaseRow'), (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $allCols = [System.Collections.Generic.List[string]]::new()
        if ($script:BaseHeaders -and $script:BaseHeaders.Count -gt 0) {
            foreach ($h in $script:BaseHeaders) { $allCols.Add($h) }
        } elseif ($sel.Record.Changes) {
            foreach ($c in $sel.Record.Changes) { $allCols.Add($c.BaseColumn) }
        }

        if ($allCols.Count -eq 0) { return }

        $p = Get-UpdaterThemePalette $script:CurrentTheme
        $editWin = New-Object System.Windows.Window -Property @{
            Title = Get-UiString 'EditBaseFieldTitle'
            Width = 480
            Height = 320
            WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner
            Owner = $window
            Background = $script:BrushConverter.ConvertFromString($p.BgCard)
            Foreground = $script:BrushConverter.ConvertFromString($p.TextPrimary)
            FontFamily = New-Object System.Windows.Media.FontFamily('Segoe UI')
            FontSize = 13
        }
        $editHwnd = (New-Object System.Windows.Interop.WindowInteropHelper($editWin)).EnsureHandle()
        Set-WindowDwmTheme -Hwnd $editHwnd -IsDark $p.IsDark
        if ($null -ne $window.Resources) {
            foreach ($k in $window.Resources.Keys) { if ($null -ne $window.Resources[$k]) { $editWin.Resources[$k] = $window.Resources[$k] } }
        }

        $sp = New-Object System.Windows.Controls.StackPanel -Property @{ Margin = New-Object System.Windows.Thickness(20) }

        # Column picker
        $lblCol = New-Object System.Windows.Controls.TextBlock -Property @{ Text = Get-UiString 'ColField'; Foreground = $script:BrushConverter.ConvertFromString($p.TextSecondary); Margin = New-Object System.Windows.Thickness(0, 0, 0, 4) }
        $cmbCol = New-Object System.Windows.Controls.ComboBox -Property @{ Background = $script:BrushConverter.ConvertFromString($p.BgInput); Foreground = $script:BrushConverter.ConvertFromString($p.TextPrimary); BorderBrush = $script:BrushConverter.ConvertFromString($p.BorderInput); Margin = New-Object System.Windows.Thickness(0, 0, 0, 10) }
        foreach ($c in $allCols) { [void]$cmbCol.Items.Add($c) }

        # Current value in base display
        $lblCur = New-Object System.Windows.Controls.TextBlock -Property @{ Text = Get-UiString 'ColCurrentBase'; Foreground = $script:BrushConverter.ConvertFromString($p.TextSecondary); Margin = New-Object System.Windows.Thickness(0, 0, 0, 4) }
        $txtCurVal = New-Object System.Windows.Controls.TextBlock -Property @{ Text = ''; Foreground = $script:BrushConverter.ConvertFromString($p.TextMuted); FontWeight = [System.Windows.FontWeights]::SemiBold; Margin = New-Object System.Windows.Thickness(2, 0, 0, 10) }

        # New Value
        $lblVal = New-Object System.Windows.Controls.TextBlock -Property @{ Text = Get-UiString 'ColIncomingValue'; Foreground = $script:BrushConverter.ConvertFromString($p.TextSecondary); Margin = New-Object System.Windows.Thickness(0, 0, 0, 4) }
        $txtVal = New-Object System.Windows.Controls.TextBox -Property @{
            Background = $script:BrushConverter.ConvertFromString($p.BgInput)
            Foreground = $script:BrushConverter.ConvertFromString($p.TextPrimary)
            BorderBrush = $script:BrushConverter.ConvertFromString($p.BorderInput)
            Padding = New-Object System.Windows.Thickness(6, 4, 6, 4)
            Margin = New-Object System.Windows.Thickness(0, 0, 0, 16)
        }

        # Helper to refresh current and initial new value for selected column
        $UpdateValuesForCol = {
            if ($allCols -and $cmbCol.SelectedIndex -ge 0 -and $cmbCol.SelectedIndex -lt $allCols.Count) {
                $targetCol = $allCols[$cmbCol.SelectedIndex]
                $valInBase = ''
                if ($sel.Record.MatchedBaseRow -and $sel.Record.MatchedBaseRow.Values) {
                    if ($sel.Record.MatchedBaseRow.Values.ContainsKey($targetCol)) {
                        $valInBase = $sel.Record.MatchedBaseRow.Values[$targetCol]
                    }
                } elseif ($sel.Record.ProjectedRow) {
                    if ($sel.Record.ProjectedRow.ContainsKey($targetCol)) {
                        $valInBase = $sel.Record.ProjectedRow[$targetCol]
                    }
                }
                $txtCurVal.Text = if ($null -ne $valInBase) { $valInBase.ToString() } else { '' }

                # Prepopulate new value: if already changed, show new value, else base value
                $chgVal = $valInBase
                if ($sel.Record.Changes) {
                    $foundChg = $sel.Record.Changes | Where-Object { $_.BaseColumn -eq $targetCol } | Select-Object -First 1
                    if ($foundChg -and $null -ne $foundChg.NewValue) {
                        $chgVal = $foundChg.NewValue
                    }
                }
                $txtVal.Text = if ($null -ne $chgVal) { $chgVal.ToString() } else { '' }
                $txtVal.SelectAll()
            }
        }

        $cmbCol.add_SelectionChanged({ & $UpdateValuesForCol })

        # Determine pre-selected column
        $selectedIdx = 0
        $currentCellCol = if ($dgBaseFullRow -and $dgBaseFullRow.CurrentColumn) { $dgBaseFullRow.CurrentColumn.Header.ToString() } else { '' }
        if (-not [string]::IsNullOrEmpty($currentCellCol) -and $allCols.Contains($currentCellCol)) {
            $selectedIdx = $allCols.IndexOf($currentCellCol)
        } elseif ($sel.Record.Changes -and $sel.Record.Changes.Count -gt 0 -and $sel.Record.Changes[0]) {
            $firstChg = $sel.Record.Changes[0].BaseColumn
            if ($allCols.Contains($firstChg)) { $selectedIdx = $allCols.IndexOf($firstChg) }
        }
        $cmbCol.SelectedIndex = $selectedIdx
        & $UpdateValuesForCol

        [void]$sp.Children.Add($lblCol); [void]$sp.Children.Add($cmbCol)
        [void]$sp.Children.Add($lblCur); [void]$sp.Children.Add($txtCurVal)
        [void]$sp.Children.Add($lblVal); [void]$sp.Children.Add($txtVal)

        # Buttons
        $btnSp = New-Object System.Windows.Controls.StackPanel -Property @{ Orientation = [System.Windows.Controls.Orientation]::Horizontal; HorizontalAlignment = [System.Windows.HorizontalAlignment]::Right }
        $btnSave = New-Object System.Windows.Controls.Button -Property @{ Content = Get-UiString 'SettingsBtnSave'; Background = $script:BrushConverter.ConvertFromString($p.AccentBlue); Foreground = $script:BrushConverter.ConvertFromString('#FFFFFF'); FontWeight = [System.Windows.FontWeights]::Bold; Padding = New-Object System.Windows.Thickness(14, 6, 14, 6); Margin = New-Object System.Windows.Thickness(0, 0, 8, 0) }
        $btnCancel = New-Object System.Windows.Controls.Button -Property @{ Content = Get-UiString 'SettingsBtnCancel'; Background = $script:BrushConverter.ConvertFromString($p.BtnSecondaryBg); Foreground = $script:BrushConverter.ConvertFromString($p.BtnSecondaryFg); BorderBrush = $script:BrushConverter.ConvertFromString($p.BorderCard); BorderThickness = [System.Windows.Thickness]::new(1); Padding = New-Object System.Windows.Thickness(14, 6, 14, 6) }

        $btnSave.add_Click({
            if ($allCols -and $cmbCol.SelectedIndex -ge 0 -and $cmbCol.SelectedIndex -lt $allCols.Count) {
                $targetCol = $allCols[$cmbCol.SelectedIndex]
                $newVal = $txtVal.Text
                $rec = $sel.Record

                $oldVal = ''
                if ($rec.MatchedBaseRow -and $rec.MatchedBaseRow.Values) {
                    $oldVal = if ($rec.MatchedBaseRow.Values.ContainsKey($targetCol)) { $rec.MatchedBaseRow.Values[$targetCol] } else { '' }
                    $rec.MatchedBaseRow.Values[$targetCol] = $newVal
                } elseif ($rec.ProjectedRow) {
                    $oldVal = if ($rec.ProjectedRow.ContainsKey($targetCol)) { $rec.ProjectedRow[$targetCol] } else { '' }
                    $rec.ProjectedRow[$targetCol] = $newVal
                }

                if ($null -eq $rec.Changes -or $rec.Changes -isnot [System.Collections.Generic.List[object]]) {
                    $rec.Changes = [System.Collections.Generic.List[object]]::new([object[]]@($rec.Changes))
                }
                $existing = $rec.Changes | Where-Object { $_.BaseColumn -eq $targetCol } | Select-Object -First 1
                if ($existing) {
                    $existing.NewValue = $newVal
                    $existing.SelectedForUpdate = $true
                } else {
                    $rec.Changes.Add([PSCustomObject]@{
                        BaseColumn        = $targetCol
                        OldValue          = $oldVal
                        NewValue          = $newVal
                        SelectedForUpdate = $true
                    })
                }
                if (-not $sel.SelectedCells) { $sel.SelectedCells = @{} }; $sel.SelectedCells[$targetCol] = $true

                if ($rec.Status -eq 'Unchanged') {
                    $rec.Status = 'Changed'
                    $sel.StatusText = Get-UiString 'StatusBadgeChanged'
                    $bColors = Get-StatusBadgeColors 'Changed' $script:CurrentTheme
                    $sel.StatusBg = $bColors.Bg
                    $sel.StatusFg = $bColors.Fg
                    & $UpdateCounters
                }

                $btnApplyAccepted.IsEnabled = ($script:AllReviewItems | Where-Object { $_.Decision -eq 'Accepted' }).Count -gt 0

                & $RenderDetailPane $sel
                & $RenderFullBaseRow $sel
            }
            $editWin.Close()
        })
        $btnCancel.add_Click({ $editWin.Close() })

        [void]$btnSp.Children.Add($btnSave); [void]$btnSp.Children.Add($btnCancel)
        [void]$sp.Children.Add($btnSp)
        $editWin.Content = $sp
        [void]$editWin.ShowDialog()
    }

    $btnEditRow.add_Click({ & $OpenModifyFieldDialog })
    $btnModifyBaseField.add_Click({ & $OpenModifyFieldDialog })
    if ($lbReviewItems) {
        $lbReviewItems.add_MouseDoubleClick({
            if ($lbReviewItems.SelectedItem) {
                & $OpenModifyFieldDialog
            }
        })
    }

    # Accept All Button
    $btnAcceptAll.add_Click({
        if (-not $script:AllReviewItems -or $script:AllReviewItems.Count -eq 0) { return }
        $count = 0
        $badge = Get-StatusBadgeColors -Status 'Accepted' -ThemeName $script:CurrentTheme
        foreach ($item in $script:AllReviewItems) {
            if ($item.Record.Status -ne 'Ambiguous' -and $item.Record.Status -ne 'Unchanged') {
                $item.Decision = 'Accepted'
                $item.StatusText = Get-UiString 'StatusBadgeAccepted'
                $item.StatusBg = $badge.Bg
                $item.StatusFg = $badge.Fg
                $count++
            }
        }
        & $UpdateCounters
        $btnApplyAccepted.IsEnabled = ($script:AllReviewItems | Where-Object { $_.Decision -eq 'Accepted' }).Count -gt 0
        [System.Windows.Forms.MessageBox]::Show(((Get-UiString 'AcceptAllDone') -f $count), (Get-UiString 'InfoTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })

    # Reject All Button
    $btnRejectAll.add_Click({
        if (-not $script:AllReviewItems -or $script:AllReviewItems.Count -eq 0) { return }
        $count = 0
        $badge = Get-StatusBadgeColors -Status 'Rejected' -ThemeName $script:CurrentTheme
        foreach ($item in $script:AllReviewItems) {
            $item.Decision = 'Rejected'
            $item.StatusText = Get-UiString 'StatusBadgeRejected'
            $item.StatusBg = $badge.Bg
            $item.StatusFg = $badge.Fg
            $count++
        }
        & $UpdateCounters
        $btnApplyAccepted.IsEnabled = ($script:AllReviewItems | Where-Object { $_.Decision -eq 'Accepted' }).Count -gt 0
        [System.Windows.Forms.MessageBox]::Show(((Get-UiString 'RejectAllDone') -f $count), (Get-UiString 'InfoTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })

    # Export Report Button
    $btnExportReport.add_Click({
        if (-not $script:AllReviewItems -or $script:AllReviewItems.Count -eq 0) { return }
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = (Get-UiString 'FilterExportReport')
        $sfd.FileName = (Get-UiString 'ExportReportFileName') -f (Get-Date -Format 'yyyyMMdd_HHmmss')
        if ($sfd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

        try {
            [void](Export-ReviewReport -ReviewItems $script:AllReviewItems -DestinationPath $sfd.FileName -BaseFilePath $txtBasePath.Text -IncomingPath $txtIncomingPath.Text)
            [System.Windows.Forms.MessageBox]::Show(((Get-UiString 'ExportSuccess') -f $sfd.FileName), (Get-UiString 'InfoTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show(((Get-UiString 'ErrExportReport') -f $_.Exception.Message), (Get-UiString 'ErrorTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    # Keyboard Shortcuts
    $window.add_KeyDown({
        param($s, $e)
        if ($e.Key -eq 'F1') {
            if ($btnHelp) { $btnHelp.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            $e.Handled = $true
            return
        }

        if ($e.OriginalSource -is [System.Windows.Controls.TextBox]) { return }

        # P3 — Ctrl+Z: undo last decision
        if ($e.Key -eq 'Z' -and [System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control) {
            & $PerformUndo
            return
        }

        switch ($e.Key) {
            'B'    { $btnBackRow.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            'Back' { $btnBackRow.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            'Left' { $btnBackRow.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            'A'    { $btnAcceptRow.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            'R'    { $btnRejectRow.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            'S'    { $btnSkipRow.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            'E'    { $btnEditRow.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
        }
    })

    # Centralized Undo Action
    $PerformUndo = {
        if ($script:UndoStack.Count -gt 0) {
            $snap = $script:UndoStack.Pop()
            $snap.Item.Decision   = $snap.PrevDecision
            $snap.Item.StatusText = $snap.PrevStatusText
            $snap.Item.StatusBg   = $snap.PrevStatusBg
            $snap.Item.StatusFg   = $snap.PrevStatusFg
            & $UpdateCounters
            $btnApplyAccepted.IsEnabled = ($script:AllReviewItems | Where-Object { $_.Decision -eq 'Accepted' }).Count -gt 0
            $txtStatusMsg.Text = (Get-UiString 'UndoStatusMsg') -f $snap.PrevDecision
        }
    }
    if ($btnUndo) {
        $btnUndo.add_Click({ & $PerformUndo })
    }

    # Run Compare Button
    $btnRunCompare.add_Click({
        $basePath = $txtBasePath.Text
        $incPath  = $txtIncomingPath.Text
        $baseSheet = if ($cmbBaseSheet.SelectedItem) { $cmbBaseSheet.SelectedItem.ToString() } else { '' }
        $incSheet  = if ($cmbIncomingSheet.SelectedItem) { $cmbIncomingSheet.SelectedItem.ToString() } else { '' }

        # 1. File existence & empty path validation
        if ([string]::IsNullOrWhiteSpace($basePath) -and [string]::IsNullOrWhiteSpace($incPath)) {
            $msg = Get-UiString 'ErrSelectFiles'
            $txtStatusMsg.Text = $msg
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        if ([string]::IsNullOrWhiteSpace($basePath) -or -not (Test-Path $basePath)) {
            $msg = Get-UiString 'ErrSelectBaseFile'
            $txtStatusMsg.Text = $msg
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        if ([string]::IsNullOrWhiteSpace($incPath) -or -not (Test-Path $incPath)) {
            $msg = Get-UiString 'ErrSelectIncomingFile'
            $txtStatusMsg.Text = $msg
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        # 2. Same file check
        if ([System.IO.Path]::GetFullPath($basePath).ToLowerInvariant() -eq [System.IO.Path]::GetFullPath($incPath).ToLowerInvariant()) {
            $msg = Get-UiString 'ErrSameFile'
            $txtStatusMsg.Text = $msg
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        # 3. File format check
        $baseExt = [System.IO.Path]::GetExtension($basePath).ToLowerInvariant()
        if ($baseExt -notin @('.xlsx', '.csv')) {
            $msg = (Get-UiString 'ErrUnsupportedFormat') -f $baseExt
            $txtStatusMsg.Text = $msg
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $incExt = [System.IO.Path]::GetExtension($incPath).ToLowerInvariant()
        if ($incExt -notin @('.xlsx', '.csv')) {
            $msg = (Get-UiString 'ErrUnsupportedFormat') -f $incExt
            $txtStatusMsg.Text = $msg
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        # 4. Join key selection validation
        $baseKeys = @($lbJoinBase.SelectedItems | ForEach-Object { $_.ToString() })
        $incKeys  = @($lbJoinIncoming.SelectedItems | ForEach-Object { $_.ToString() })
        if ($baseKeys.Length -eq 0 -or $incKeys.Length -eq 0) {
            $msg = Get-UiString 'ErrSelectJoinKeys'
            $txtStatusMsg.Text = $msg
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        if ($baseKeys.Length -ne $incKeys.Length) {
            $msg = (Get-UiString 'ErrJoinKeyCountMismatch') -f $baseKeys.Length, ($baseKeys -join ', '), $incKeys.Length, ($incKeys -join ', ')
            $txtStatusMsg.Text = $msg
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        # 5. Mapping rules check
        if ($script:MappingRules.Count -eq 0) {
            $msg = Get-UiString 'ErrDefineRules'
            $txtStatusMsg.Text = $msg
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $txtStatusMsg.Text = Get-UiString 'StatusComparing'
        [System.Windows.Forms.Application]::DoEvents()

        # 6. Safe ReadSheet with error trapping
        try {
            $baseRows = [FastExcelHelper]::ReadSheet($basePath, $baseSheet)
            $incRows  = [FastExcelHelper]::ReadSheet($incPath, $incSheet)
            # Stamp SourceFilePath on each incoming row so metadata tokens SourceFileFullPath / SourceFileName work
            $incFullPath = [System.IO.Path]::GetFullPath($incPath)
            foreach ($r in $incRows) {
                if (-not $r.PSObject.Properties['SourceFilePath']) {
                    $r | Add-Member -NotePropertyName 'SourceFilePath' -NotePropertyValue $incFullPath -Force
                } elseif ([string]::IsNullOrEmpty($r.SourceFilePath)) {
                    $r.SourceFilePath = $incFullPath
                }
            }
        } catch {
            $msg = (Get-UiString 'ErrFileLockedOrCorrupt') -f [System.IO.Path]::GetFileName($incPath), $_.Exception.Message
            $txtStatusMsg.Text = (Get-UiString 'ErrReadData') -f $_.Exception.Message
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'ErrorTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        if (-not $incRows -or $incRows.Count -eq 0) {
            $msg = Get-UiString 'ErrNoDataRows'
            $txtStatusMsg.Text = $msg
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $compOptions = @{
            IgnoreCase         = if ($script:chkIgnoreCase) { [bool]$script:chkIgnoreCase.IsChecked } else { $true }
            Trim               = if ($script:chkTrimWhitespace) { [bool]$script:chkTrimWhitespace.IsChecked } else { $true }
            TrimWhitespace     = if ($script:chkTrimWhitespace) { [bool]$script:chkTrimWhitespace.IsChecked } else { $true }
            IgnoreSpecialChars = if ($script:chkIgnoreSpecialChars) { [bool]$script:chkIgnoreSpecialChars.IsChecked } else { $true }
            IgnoreAllSpaces    = if ($script:chkIgnoreAllSpaces) { [bool]$script:chkIgnoreAllSpaces.IsChecked } else { $true }
        }

        $comp = Invoke-MasterCompare -BaseRows $baseRows -IncomingRows $incRows -MappingProfile $script:CurrentProfile -MappingRules $script:MappingRules -BaseJoinKey $baseKeys -IncomingJoinKey $incKeys -CompareOptions $compOptions -BaseHeaders $script:BaseHeaders -DetectRemoved ($script:AppConfig.DetectRemovedRows -eq $true) -MarkDeletedColumn $script:AppConfig.MarkDeletedColumn -MarkDeletedValue $script:AppConfig.MarkDeletedValue
        $script:ComparisonResult = $comp

        & $script:PopulateReviewItems
        if ($script:mainTabs) { $script:mainTabs.SelectedIndex = 1 }
        elseif ($mainTabs) { $mainTabs.SelectedIndex = 1 }
        if ($script:txtStatusMsg) { $script:txtStatusMsg.Text = (Get-UiString 'StatusCompareDone') -f $script:AllReviewItems.Count }
        elseif ($txtStatusMsg) { $txtStatusMsg.Text = (Get-UiString 'StatusCompareDone') -f $script:AllReviewItems.Count }
    })

    # MainTabs Selection Changed (Refresh preview on tab 1)
    if ($mainTabs) {
        $mainTabs.add_SelectionChanged({
            param($s, $e)
            if ($e.Source -eq $mainTabs -and $mainTabs.SelectedItem -eq $tabMapping) {
                & $script:UpdateDataMappingPreview
            }
        })
    }

    # Apply Accepted Changes Button
    $btnApplyAccepted.add_Click({
        $accepted = @($script:AllReviewItems | Where-Object { $_.Decision -eq 'Accepted' })
        if ($accepted.Length -eq 0) {
            [System.Windows.Forms.MessageBox]::Show((Get-UiString 'StatusNoAccepted'), (Get-UiString 'InfoTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }

        $confirmMsg = (Get-UiString 'ConfirmApplyMsg') -f $accepted.Length
        $res = [System.Windows.Forms.MessageBox]::Show($confirmMsg, (Get-UiString 'ConfirmApplyTitle'), [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($res -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $txtStatusMsg.Text = Get-UiString 'StatusWriting'
        [System.Windows.Forms.Application]::DoEvents()

        # Synchronize selective cell updates
        $accRecords = [System.Collections.Generic.List[object]]::new()
        foreach ($item in $accepted) {
            $rec = $item.Record
            if ($rec.Changes) {
                foreach ($chg in $rec.Changes) {
                    $chg.SelectedForUpdate = if ($item.SelectedCells) { ($item.SelectedCells[$chg.BaseColumn] -ne $false) } else { $true }
                }
            }
            $accRecords.Add($rec)
        }

        $basePath  = $txtBasePath.Text
        $baseSheet = if ($cmbBaseSheet.SelectedItem) { $cmbBaseSheet.SelectedItem.ToString() } else { '' }

        try {
            $wbRes = Invoke-MasterWriteBack -BaseFilePath $basePath -BaseSheet $baseSheet -AcceptedItems $accRecords -AppConfig $script:AppConfig
            $logRes = Write-ImportLog -LogDirectory $script:AppConfig.LogDirectory -BatchId $wbRes.BatchId -ReviewItems $script:AllReviewItems -WriteBackResult $wbRes -BaseFilePath $basePath -RedactNames $script:AppConfig.RedactNamesInLog -LogChangesToBaseSheet ($script:AppConfig.LogChangesToBaseSheet -eq $true) -BaseSheetLogName ($(if ($script:AppConfig.BaseSheetLogName) { $script:AppConfig.BaseSheetLogName } else { 'ImportLog' }))

            $succMsg = (Get-UiString 'MsgWriteSuccessSummary') -f (Get-UiString 'StatusWriteSuccess'), $wbRes.UpdatedCells, $wbRes.AddedRows, $wbRes.BackupPath, $logRes.TxtPath
            if ($logRes.BaseSheetLog) {
                $succMsg += ((Get-UiString 'MsgBaseSheetLogSuccess') -f $logRes.BaseSheetLog)
            }
            # P6: Surface backup-size-prune warning
            if ($wbRes.BackupSizeWarning) {
                $succMsg += "`n`n" + ((Get-UiString 'MsgBackupSizeWarning') -f $wbRes.BackupSizeWarning)
            }
            [System.Windows.Forms.MessageBox]::Show($succMsg, (Get-UiString 'InfoTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $txtStatusMsg.Text = Get-UiString 'StatusWriteSuccess'
            $btnApplyAccepted.IsEnabled = $false
            if ($btnApplyToNewFile) { $btnApplyToNewFile.IsEnabled = $false }
        } catch {
                        $msg = (Get-UiString 'ErrSaveWriteback') -f $_.Exception.Message
            [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'ErrorTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $txtStatusMsg.Text = (Get-UiString 'StatusWriteError') -f $_.Message
        }
    })

    # Apply to New File Button (Create a new updated copy without touching original base file)
    if ($btnApplyToNewFile) {
        $btnApplyToNewFile.add_Click({
            $accepted = @($script:AllReviewItems | Where-Object { $_.Decision -eq 'Accepted' })
            if ($accepted.Length -eq 0) {
                [System.Windows.Forms.MessageBox]::Show((Get-UiString 'StatusNoAccepted'), (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }

            $basePath  = $txtBasePath.Text
            $baseSheet = if ($cmbBaseSheet.SelectedItem) { $cmbBaseSheet.SelectedItem.ToString() } else { '' }
            if (-not (Test-Path $basePath)) {
                $errFileNotFound = (Get-UiString 'ErrSaveWriteback') -f "Base file not found: $basePath"
                [System.Windows.Forms.MessageBox]::Show($errFileNotFound, (Get-UiString 'ErrorTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }

            $baseExt = [System.IO.Path]::GetExtension($basePath)
            $baseDir = [System.IO.Path]::GetDirectoryName($basePath)
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($basePath)
            $defaultNewName = "{0}_Updated_{1}{2}" -f $baseName, (Get-Date -Format 'yyyyMMdd_HHmm'), $baseExt

            $sfd = New-Object System.Windows.Forms.SaveFileDialog
            $sfd.Title = Get-UiString 'BtnApplyToNewFile'
            if (Test-Path $baseDir) { $sfd.InitialDirectory = $baseDir }
            $sfd.FileName = $defaultNewName
            $sfd.Filter = if ($baseExt -eq '.csv') { "CSV (*.csv)|*.csv|All Files (*.*)|*.*" } else { "Excel Workbook (*.xlsx)|*.xlsx|All Files (*.*)|*.*" }

            if ($sfd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
            $newFilePath = $sfd.FileName

            # Check if user selected the exact same file path as the original base file
            if ([string]::Equals([System.IO.Path]::GetFullPath($newFilePath), [System.IO.Path]::GetFullPath($basePath), [System.StringComparison]::OrdinalIgnoreCase)) {
                $sameFileWarn = if ($script:CurrentLanguage -eq 'pl') {
                    "Wybrano ten sam plik bazy! Użyj przycisku 'Zastosuj do bazy' lub wybierz inną nazwę nowego pliku."
                } elseif ($script:CurrentLanguage -eq 'de') {
                    "Sie haben dieselbe Datei ausgewählt! Verwenden Sie 'Anwenden' oder wählen Sie einen anderen Dateinamen."
                } else {
                    "You selected the original base file! Use 'Apply to Base' or specify a different new file name."
                }
                [System.Windows.Forms.MessageBox]::Show($sameFileWarn, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }

            $txtStatusMsg.Text = Get-UiString 'StatusWriting'
            [System.Windows.Forms.Application]::DoEvents()

            # Synchronize selective cell updates
            $accRecords = [System.Collections.Generic.List[object]]::new()
            foreach ($item in $accepted) {
                $rec = $item.Record
                if ($rec.Changes) {
                    foreach ($chg in $rec.Changes) {
                        $chg.SelectedForUpdate = if ($item.SelectedCells) { ($item.SelectedCells[$chg.BaseColumn] -ne $false) } else { $true }
                    }
                }
                $accRecords.Add($rec)
            }

            try {
                # Copy original base file to new destination
                [System.IO.File]::Copy($basePath, $newFilePath, $true)

                # Write changes into the new file
                $wbRes = Invoke-MasterWriteBack -BaseFilePath $newFilePath -BaseSheet $baseSheet -AcceptedItems $accRecords -AppConfig $script:AppConfig
                $logRes = Write-ImportLog -LogDirectory $script:AppConfig.LogDirectory -BatchId $wbRes.BatchId -ReviewItems $script:AllReviewItems -WriteBackResult $wbRes -BaseFilePath $newFilePath -RedactNames $script:AppConfig.RedactNamesInLog -LogChangesToBaseSheet ($script:AppConfig.LogChangesToBaseSheet -eq $true) -BaseSheetLogName ($(if ($script:AppConfig.BaseSheetLogName) { $script:AppConfig.BaseSheetLogName } else { 'ImportLog' }))

                $succMsg = (Get-UiString 'MsgWriteNewFileSuccess') -f $newFilePath, $wbRes.UpdatedCells, $wbRes.AddedRows, $logRes.TxtPath
                [System.Windows.Forms.MessageBox]::Show($succMsg, (Get-UiString 'InfoTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $txtStatusMsg.Text = Get-UiString 'StatusWriteSuccess'
            } catch {
                $msg = (Get-UiString 'ErrSaveWriteback') -f $_.Exception.Message
                [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'ErrorTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $txtStatusMsg.Text = (Get-UiString 'StatusWriteError') -f $_.Message
            }
        })
    }

    # Settings Button (Tabbed Dialog: General + Metadata Columns)
    $btnSettings.add_Click({
        $p = Get-UpdaterThemePalette $script:CurrentTheme
        $setWin = New-Object System.Windows.Window -Property @{
            Title = Get-UiString 'SettingsTitle'
            Width = 580
            Height = 580
            WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner
            Owner = $window
            Background = $script:BrushConverter.ConvertFromString($p.BgCard)
            Foreground = $script:BrushConverter.ConvertFromString($p.TextPrimary)
            FontFamily = New-Object System.Windows.Media.FontFamily('Segoe UI')
            FontSize = 13
        }
        $setHwnd = (New-Object System.Windows.Interop.WindowInteropHelper($setWin)).EnsureHandle()
        Set-WindowDwmTheme -Hwnd $setHwnd -IsDark $p.IsDark
        if ($null -ne $window.Resources) {
            foreach ($k in $window.Resources.Keys) { if ($null -ne $window.Resources[$k]) { $setWin.Resources[$k] = $window.Resources[$k] } }
        }

        $gridMain = New-Object System.Windows.Controls.Grid
        $row0 = New-Object System.Windows.Controls.RowDefinition -Property @{ Height = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star) }
        $row1 = New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto }
        [void]$gridMain.RowDefinitions.Add($row0); [void]$gridMain.RowDefinitions.Add($row1)

        $tc = New-Object System.Windows.Controls.TabControl -Property @{ Background = $script:BrushConverter.ConvertFromString($p.BgCard); BorderBrush = $script:BrushConverter.ConvertFromString($p.BorderCard); BorderThickness = New-Object System.Windows.Thickness(0, 1, 0, 0) }

        # --- TAB 1: General ---
        $tiGen = New-Object System.Windows.Controls.TabItem -Property @{ Header = Get-UiString 'SettingsTabGeneral'; Foreground = $script:BrushConverter.ConvertFromString($p.TextPrimary) }
        $spGen = New-Object System.Windows.Controls.StackPanel -Property @{ Margin = New-Object System.Windows.Thickness(16) }

        # Retention count
        $lblRet = New-Object System.Windows.Controls.TextBlock -Property @{ Text = Get-UiString 'SettingsBackupCount'; Foreground = $script:BrushConverter.ConvertFromString($p.TextSecondary); Margin = New-Object System.Windows.Thickness(0, 0, 0, 4) }
        $txtRet = New-Object System.Windows.Controls.TextBox -Property @{ Text = $script:AppConfig.BackupRetentionCount.ToString(); Background = $script:BrushConverter.ConvertFromString($p.BgInput); Foreground = $script:BrushConverter.ConvertFromString($p.TextPrimary); BorderBrush = $script:BrushConverter.ConvertFromString($p.BorderInput); Padding = New-Object System.Windows.Thickness(6, 4, 6, 4); Margin = New-Object System.Windows.Thickness(0, 0, 0, 12) }
        [void]$spGen.Children.Add($lblRet); [void]$spGen.Children.Add($txtRet)

        # Write mode
        $lblWm = New-Object System.Windows.Controls.TextBlock -Property @{ Text = Get-UiString 'SettingsWriteMode'; Foreground = $script:BrushConverter.ConvertFromString($p.TextSecondary); Margin = New-Object System.Windows.Thickness(0, 0, 0, 4) }
        $cmbWm = New-Object System.Windows.Controls.ComboBox -Property @{ Background = $script:BrushConverter.ConvertFromString($p.BgInput); Foreground = $script:BrushConverter.ConvertFromString($p.TextPrimary); BorderBrush = $script:BrushConverter.ConvertFromString($p.BorderInput); Margin = New-Object System.Windows.Thickness(0, 0, 0, 12) }
        [void]$cmbWm.Items.Add('InPlace'); [void]$cmbWm.Items.Add('SafeRewrite')
        $cmbWm.SelectedItem = $script:AppConfig.WriteMode
        [void]$spGen.Children.Add($lblWm); [void]$spGen.Children.Add($cmbWm)

        # Mask PII
        $chkPii = New-Object System.Windows.Controls.CheckBox -Property @{ Content = Get-UiString 'SettingsMaskPii'; Foreground = $script:BrushConverter.ConvertFromString($p.TextPrimary); IsChecked = ($script:AppConfig.RedactNamesInLog -eq $true); Margin = New-Object System.Windows.Thickness(0, 0, 0, 12) }
        [void]$spGen.Children.Add($chkPii)

        # Detect removed rows
        $chkRem = New-Object System.Windows.Controls.CheckBox -Property @{ Content = Get-UiString 'SettingsDetectRemoved'; Foreground = $script:BrushConverter.ConvertFromString($p.TextPrimary); IsChecked = ($script:AppConfig.DetectRemovedRows -eq $true); Margin = New-Object System.Windows.Thickness(0, 0, 0, 12) }
        [void]$spGen.Children.Add($chkRem)

        # Show unchanged rows
        $chkUnchanged = New-Object System.Windows.Controls.CheckBox -Property @{ Content = Get-UiString 'SettingsShowUnchanged'; Foreground = $script:BrushConverter.ConvertFromString($p.TextPrimary); IsChecked = ($script:AppConfig.ShowUnchangedRows -eq $true -or $script:AppConfig.AutoSkipUnchanged -eq $false); Margin = New-Object System.Windows.Thickness(0, 0, 0, 12) }
        [void]$spGen.Children.Add($chkUnchanged)

        # Log to base file sheet
        $chkLogBaseSheet = New-Object System.Windows.Controls.CheckBox -Property @{ Content = Get-UiString 'SettingsLogToBaseSheet'; Foreground = $script:BrushConverter.ConvertFromString($p.TextPrimary); IsChecked = ($script:AppConfig.LogChangesToBaseSheet -eq $true); Margin = New-Object System.Windows.Thickness(0, 0, 0, 6) }
        [void]$spGen.Children.Add($chkLogBaseSheet)

        $lblLogSheet = New-Object System.Windows.Controls.TextBlock -Property @{ Text = Get-UiString 'SettingsLogSheetName'; Foreground = $script:BrushConverter.ConvertFromString($p.TextSecondary); Margin = New-Object System.Windows.Thickness(20, 0, 0, 4) }
        $txtLogSheet = New-Object System.Windows.Controls.TextBox -Property @{
            Text        = if ($script:AppConfig.BaseSheetLogName) { $script:AppConfig.BaseSheetLogName } else { 'ImportLog' }
            IsEnabled   = ($script:AppConfig.LogChangesToBaseSheet -eq $true)
            Background  = $script:BrushConverter.ConvertFromString($p.BgInput)
            Foreground  = $script:BrushConverter.ConvertFromString($p.TextPrimary)
            BorderBrush = $script:BrushConverter.ConvertFromString($p.BorderInput)
            Padding     = New-Object System.Windows.Thickness(6, 4, 6, 4)
            Margin      = New-Object System.Windows.Thickness(20, 0, 0, 12)
        }
        $chkLogBaseSheet.add_Checked({ $txtLogSheet.IsEnabled = $true })
        $chkLogBaseSheet.add_Unchecked({ $txtLogSheet.IsEnabled = $false })
        [void]$spGen.Children.Add($lblLogSheet); [void]$spGen.Children.Add($txtLogSheet)

        # Remember base file path checkbox
        $chkRemBase = New-Object System.Windows.Controls.CheckBox -Property @{ Content = Get-UiString 'SettingsRememberBasePath'; Foreground = $script:BrushConverter.ConvertFromString($p.TextPrimary); IsChecked = ($script:AppConfig.RememberBasePath -ne $false); Margin = New-Object System.Windows.Thickness(0, 0, 0, 10) }
        [void]$spGen.Children.Add($chkRemBase)

        # Base file path textbox + browse
        $lblBaseFile = New-Object System.Windows.Controls.TextBlock -Property @{ Text = Get-UiString 'SettingsBaseFilePath'; Foreground = $script:BrushConverter.ConvertFromString($p.TextSecondary); Margin = New-Object System.Windows.Thickness(0, 0, 0, 4) }
        $gridBaseFile = New-Object System.Windows.Controls.Grid -Property @{ Margin = New-Object System.Windows.Thickness(0, 0, 0, 12) }
        $colB0 = New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star) }
        $colB1 = New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::Auto }
        [void]$gridBaseFile.ColumnDefinitions.Add($colB0); [void]$gridBaseFile.ColumnDefinitions.Add($colB1)

        $txtSetBase = New-Object System.Windows.Controls.TextBox -Property @{
            Text        = if ($script:AppConfig.BaseFilePath) { $script:AppConfig.BaseFilePath } else { '' }
            Background  = $script:BrushConverter.ConvertFromString($p.BgInput)
            Foreground  = $script:BrushConverter.ConvertFromString($p.TextPrimary)
            BorderBrush = $script:BrushConverter.ConvertFromString($p.BorderInput)
            Padding     = New-Object System.Windows.Thickness(6, 4, 6, 4)
        }
        $btnBrowseSetBase = New-Object System.Windows.Controls.Button -Property @{
            Content         = Get-UiString 'SettingsBtnBrowseBase'
            Background      = $script:BrushConverter.ConvertFromString($p.BtnSecondaryBg)
            Foreground      = $script:BrushConverter.ConvertFromString($p.BtnSecondaryFg)
            BorderBrush     = $script:BrushConverter.ConvertFromString($p.BorderCard)
            BorderThickness = [System.Windows.Thickness]::new(1)
            Padding         = New-Object System.Windows.Thickness(10, 4, 10, 4)
            Margin          = New-Object System.Windows.Thickness(6, 0, 0, 0)
        }
        $btnBrowseSetBase.add_Click({
            $ofd = New-Object System.Windows.Forms.OpenFileDialog
            $ofd.Filter = (Get-UiString 'FilterExcelCsv')
            if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $txtSetBase.Text = $ofd.FileName
            }
        })
        [System.Windows.Controls.Grid]::SetColumn($txtSetBase, 0)
        [System.Windows.Controls.Grid]::SetColumn($btnBrowseSetBase, 1)
        [void]$gridBaseFile.Children.Add($txtSetBase)
        [void]$gridBaseFile.Children.Add($btnBrowseSetBase)
        [void]$spGen.Children.Add($lblBaseFile)
        [void]$spGen.Children.Add($gridBaseFile)

        # Backups path & open button
        $lblSetBackups = New-Object System.Windows.Controls.TextBlock -Property @{ Text = Get-UiString 'SettingsBackupDir'; Foreground = $script:BrushConverter.ConvertFromString($p.TextSecondary); Margin = New-Object System.Windows.Thickness(0, 0, 0, 4) }
        $gridSetBackups = New-Object System.Windows.Controls.Grid -Property @{ Margin = New-Object System.Windows.Thickness(0, 0, 0, 10) }
        $colBk0 = New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star) }
        $colBk1 = New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::Auto }
        [void]$gridSetBackups.ColumnDefinitions.Add($colBk0); [void]$gridSetBackups.ColumnDefinitions.Add($colBk1)
        $txtSetBackups = New-Object System.Windows.Controls.TextBox -Property @{
            Text        = $script:AppConfig.BackupDirectory
            IsReadOnly  = $true
            Background  = $script:BrushConverter.ConvertFromString($p.BgInput)
            Foreground  = $script:BrushConverter.ConvertFromString($p.TextPrimary)
            BorderBrush = $script:BrushConverter.ConvertFromString($p.BorderInput)
            Padding     = New-Object System.Windows.Thickness(6, 4, 6, 4)
        }
        $btnOpenSetBackups = New-Object System.Windows.Controls.Button -Property @{
            Content         = Get-UiString 'SettingsBtnOpenBackups'
            Background      = $script:BrushConverter.ConvertFromString($p.BtnSecondaryBg)
            Foreground      = $script:BrushConverter.ConvertFromString($p.BtnSecondaryFg)
            BorderBrush     = $script:BrushConverter.ConvertFromString($p.BorderCard)
            BorderThickness = [System.Windows.Thickness]::new(1)
            Padding         = New-Object System.Windows.Thickness(10, 4, 10, 4)
            Margin          = New-Object System.Windows.Thickness(6, 0, 0, 0)
        }
        $btnOpenSetBackups.add_Click({
            $bDir = $script:AppConfig.BackupDirectory
            if (-not (Test-Path $bDir)) { [void][System.IO.Directory]::CreateDirectory($bDir) }
            [System.Diagnostics.Process]::Start('explorer.exe', $bDir) | Out-Null
        })
        [System.Windows.Controls.Grid]::SetColumn($txtSetBackups, 0)
        [System.Windows.Controls.Grid]::SetColumn($btnOpenSetBackups, 1)
        [void]$gridSetBackups.Children.Add($txtSetBackups)
        [void]$gridSetBackups.Children.Add($btnOpenSetBackups)
        [void]$spGen.Children.Add($lblSetBackups)
        [void]$spGen.Children.Add($gridSetBackups)

        # Logs path & open button
        $lblSetLogs = New-Object System.Windows.Controls.TextBlock -Property @{ Text = Get-UiString 'SettingsLogDir'; Foreground = $script:BrushConverter.ConvertFromString($p.TextSecondary); Margin = New-Object System.Windows.Thickness(0, 0, 0, 4) }
        $gridSetLogs = New-Object System.Windows.Controls.Grid -Property @{ Margin = New-Object System.Windows.Thickness(0, 0, 0, 10) }
        $colLg0 = New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star) }
        $colLg1 = New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::Auto }
        [void]$gridSetLogs.ColumnDefinitions.Add($colLg0); [void]$gridSetLogs.ColumnDefinitions.Add($colLg1)
        $txtSetLogs = New-Object System.Windows.Controls.TextBox -Property @{
            Text        = $script:AppConfig.LogDirectory
            IsReadOnly  = $true
            Background  = $script:BrushConverter.ConvertFromString($p.BgInput)
            Foreground  = $script:BrushConverter.ConvertFromString($p.TextPrimary)
            BorderBrush = $script:BrushConverter.ConvertFromString($p.BorderInput)
            Padding     = New-Object System.Windows.Thickness(6, 4, 6, 4)
        }
        $btnOpenSetLogs = New-Object System.Windows.Controls.Button -Property @{
            Content         = Get-UiString 'SettingsBtnOpenLogs'
            Background      = $script:BrushConverter.ConvertFromString($p.BtnSecondaryBg)
            Foreground      = $script:BrushConverter.ConvertFromString($p.BtnSecondaryFg)
            BorderBrush     = $script:BrushConverter.ConvertFromString($p.BorderCard)
            BorderThickness = [System.Windows.Thickness]::new(1)
            Padding         = New-Object System.Windows.Thickness(10, 4, 10, 4)
            Margin          = New-Object System.Windows.Thickness(6, 0, 0, 0)
        }
        $btnOpenSetLogs.add_Click({
            $lDir = $script:AppConfig.LogDirectory
            if (-not (Test-Path $lDir)) { [void][System.IO.Directory]::CreateDirectory($lDir) }
            [System.Diagnostics.Process]::Start('explorer.exe', $lDir) | Out-Null
        })
        [System.Windows.Controls.Grid]::SetColumn($txtSetLogs, 0)
        [System.Windows.Controls.Grid]::SetColumn($btnOpenSetLogs, 1)
        [void]$gridSetLogs.Children.Add($txtSetLogs)
        [void]$gridSetLogs.Children.Add($btnOpenSetLogs)
        [void]$spGen.Children.Add($lblSetLogs)
        [void]$spGen.Children.Add($gridSetLogs)

        $svGen = New-Object System.Windows.Controls.ScrollViewer -Property @{ VerticalScrollBarVisibility = 'Auto' }
        $svGen.Content = $spGen
        $tiGen.Content = $svGen
        [void]$tc.Items.Add($tiGen)

        # --- TAB 2: Metadata Columns ---
        $tiMeta = New-Object System.Windows.Controls.TabItem -Property @{ Header = Get-UiString 'SettingsTabMetadata'; Foreground = $script:BrushConverter.ConvertFromString($p.TextPrimary) }
        $gridMeta = New-Object System.Windows.Controls.Grid -Property @{ Margin = New-Object System.Windows.Thickness(16) }
        $gmRow0 = New-Object System.Windows.Controls.RowDefinition -Property @{ Height = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star) }
        $gmRow1 = New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto }
        [void]$gridMeta.RowDefinitions.Add($gmRow0); [void]$gridMeta.RowDefinitions.Add($gmRow1)

        $dgMeta = New-Object System.Windows.Controls.DataGrid -Property @{
            AutoGenerateColumns      = $false
            CanUserAddRows           = $false
            Background               = $script:BrushConverter.ConvertFromString($p.BgCard)
            Foreground               = $script:BrushConverter.ConvertFromString($p.TextPrimary)
            RowBackground            = $script:BrushConverter.ConvertFromString($p.DataGridRowBg)
            AlternatingRowBackground = $script:BrushConverter.ConvertFromString($p.DataGridAltRowBg)
            GridLinesVisibility      = [System.Windows.Controls.DataGridGridLinesVisibility]::Horizontal
            HorizontalGridLinesBrush = $script:BrushConverter.ConvertFromString($p.GridLines)
            BorderBrush              = $script:BrushConverter.ConvertFromString($p.BorderCard)
            HeadersVisibility        = [System.Windows.Controls.DataGridHeadersVisibility]::Column
            Margin                   = New-Object System.Windows.Thickness(0, 0, 0, 8)
        }

        $colBName = New-Object System.Windows.Controls.DataGridTextColumn -Property @{
            Header  = Get-UiString 'ColMetadataBase'
            Binding = New-Object System.Windows.Data.Binding('BaseColumn')
            Width   = New-Object System.Windows.Controls.DataGridLength(140)
        }
        $colToken = New-Object System.Windows.Controls.DataGridTextColumn -Property @{
            Header  = Get-UiString 'ColMetadataToken'
            Binding = New-Object System.Windows.Data.Binding('Token')
            Width   = New-Object System.Windows.Controls.DataGridLength(160)
        }
        $colFmt = New-Object System.Windows.Controls.DataGridTextColumn -Property @{
            Header  = Get-UiString 'ColMetadataFormat'
            Binding = New-Object System.Windows.Data.Binding('Format')
            Width   = New-Object System.Windows.Controls.DataGridLength(1, [System.Windows.Controls.DataGridLengthUnitType]::Star)
        }
        [void]$dgMeta.Columns.Add($colBName)
        [void]$dgMeta.Columns.Add($colToken)
        [void]$dgMeta.Columns.Add($colFmt)

        $metaItems = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
        foreach ($m in $script:AppConfig.MetadataColumns) {
            $metaItems.Add([PSCustomObject]@{
                BaseColumn = $m.BaseColumn
                Token      = $m.Token
                Format     = $m.Format
            })
        }
        $dgMeta.ItemsSource = $metaItems
        [System.Windows.Controls.Grid]::SetRow($dgMeta, 0)
        [void]$gridMeta.Children.Add($dgMeta)

        $spMetaBtns = New-Object System.Windows.Controls.StackPanel -Property @{ Orientation = [System.Windows.Controls.Orientation]::Horizontal; Margin = New-Object System.Windows.Thickness(0, 4, 0, 0) }
        $btnAddM = New-Object System.Windows.Controls.Button -Property @{ Content = Get-UiString 'BtnAddMeta'; Background = $script:BrushConverter.ConvertFromString($p.BtnSecondaryBg); Foreground = $script:BrushConverter.ConvertFromString($p.BtnSecondaryFg); BorderBrush = $script:BrushConverter.ConvertFromString($p.BorderCard); BorderThickness = [System.Windows.Thickness]::new(1); Padding = New-Object System.Windows.Thickness(10, 4, 10, 4); Margin = New-Object System.Windows.Thickness(0, 0, 8, 0) }
        $btnRemM = New-Object System.Windows.Controls.Button -Property @{ Content = Get-UiString 'BtnRemoveMeta'; Background = $script:BrushConverter.ConvertFromString($p.BtnSecondaryBg); Foreground = $script:BrushConverter.ConvertFromString($p.BtnSecondaryFg); BorderBrush = $script:BrushConverter.ConvertFromString($p.BorderCard); BorderThickness = [System.Windows.Thickness]::new(1); Padding = New-Object System.Windows.Thickness(10, 4, 10, 4) }
        $btnAddM.add_Click({
            $metaItems.Add([PSCustomObject]@{ BaseColumn = 'NowaKolumna'; Token = 'ChangeDate'; Format = 'yyyy-MM-dd HH:mm' })
        })
        $btnRemM.add_Click({
            if ($dgMeta.SelectedItem) {
                [void]$metaItems.Remove($dgMeta.SelectedItem)
            }
        })
        [void]$spMetaBtns.Children.Add($btnAddM); [void]$spMetaBtns.Children.Add($btnRemM)
        [System.Windows.Controls.Grid]::SetRow($spMetaBtns, 1)
        [void]$gridMeta.Children.Add($spMetaBtns)
        $tiMeta.Content = $gridMeta
        [void]$tc.Items.Add($tiMeta)

        [System.Windows.Controls.Grid]::SetRow($tc, 0)
        [void]$gridMain.Children.Add($tc)

        # Dialog Buttons (Save / Cancel)
        $btnSp = New-Object System.Windows.Controls.StackPanel -Property @{ Orientation = [System.Windows.Controls.Orientation]::Horizontal; HorizontalAlignment = [System.Windows.HorizontalAlignment]::Right; Margin = New-Object System.Windows.Thickness(16) }
        $btnSave = New-Object System.Windows.Controls.Button -Property @{ Content = Get-UiString 'SettingsBtnSave'; Background = $script:BrushConverter.ConvertFromString($p.AccentBlue); Foreground = $script:BrushConverter.ConvertFromString('#FFFFFF'); FontWeight = [System.Windows.FontWeights]::Bold; Padding = New-Object System.Windows.Thickness(14, 6, 14, 6); Margin = New-Object System.Windows.Thickness(0, 0, 8, 0) }
        $btnCancel = New-Object System.Windows.Controls.Button -Property @{ Content = Get-UiString 'SettingsBtnCancel'; Background = $script:BrushConverter.ConvertFromString($p.BtnSecondaryBg); Foreground = $script:BrushConverter.ConvertFromString($p.BtnSecondaryFg); BorderBrush = $script:BrushConverter.ConvertFromString($p.BorderCard); BorderThickness = [System.Windows.Thickness]::new(1); Padding = New-Object System.Windows.Thickness(14, 6, 14, 6) }

        $btnSave.add_Click({
            if ($txtRet.Text -match '^\d+$') { $script:AppConfig.BackupRetentionCount = [int]$txtRet.Text }
            if ($cmbWm.SelectedItem) { $script:AppConfig.WriteMode = $cmbWm.SelectedItem.ToString() }
            $script:AppConfig.RedactNamesInLog = ($chkPii.IsChecked -eq $true)
            $script:AppConfig.DetectRemovedRows = ($chkRem.IsChecked -eq $true)
            $script:AppConfig.ShowUnchangedRows = ($chkUnchanged.IsChecked -eq $true)
            $script:AppConfig.AutoSkipUnchanged = ($chkUnchanged.IsChecked -ne $true)
            if ($chkShowUnchanged) { $chkShowUnchanged.IsChecked = ($chkUnchanged.IsChecked -eq $true) }
            if ($script:ComparisonResult -and $script:PopulateReviewItems) { & $script:PopulateReviewItems }
            $script:AppConfig.RememberBasePath = ($chkRemBase.IsChecked -eq $true)
            $newBasePath = $txtSetBase.Text.Trim().Trim('"').Trim("'")
            $script:AppConfig.BaseFilePath = $newBasePath
            if ($chkRemBase.IsChecked -eq $false -and [string]::IsNullOrEmpty($newBasePath)) {
                $script:AppConfig.BaseFilePath = ''
                $script:AppConfig.BaseSheet = ''
            }
            if (-not [string]::IsNullOrEmpty($newBasePath) -and (Test-Path $newBasePath) -and $txtBasePath.Text -ne $newBasePath) {
                & $LoadBaseFile $newBasePath
            }

            $script:AppConfig.LogChangesToBaseSheet = ($chkLogBaseSheet.IsChecked -eq $true)
            $cleanLogSheet = if ($txtLogSheet.Text) { [regex]::Replace($txtLogSheet.Text.Trim(), '[\\/\?\*:[\]]', '_') } else { '' }
            if ([string]::IsNullOrWhiteSpace($cleanLogSheet)) { $cleanLogSheet = 'ImportLog' }
            if ($cleanLogSheet.Length -gt 31) { $cleanLogSheet = $cleanLogSheet.Substring(0, 31) }
            $script:AppConfig.BaseSheetLogName = $cleanLogSheet

            # Update MetadataColumns from DataGrid
            $newMeta = @()
            foreach ($row in $metaItems) {
                if (-not [string]::IsNullOrWhiteSpace($row.BaseColumn) -and -not [string]::IsNullOrWhiteSpace($row.Token)) {
                    $newMeta += [ordered]@{
                        BaseColumn = $row.BaseColumn.Trim()
                        Token      = $row.Token.Trim()
                        Format     = if ($row.Format) { $row.Format.Trim() } else { '' }
                    }
                }
            }
            if ($newMeta.Count -gt 0) {
                $script:AppConfig.MetadataColumns = $newMeta
            }

            Save-AppConfig -Config $script:AppConfig
            $setWin.Close()
        })
        $btnCancel.add_Click({ $setWin.Close() })

        [void]$btnSp.Children.Add($btnSave); [void]$btnSp.Children.Add($btnCancel)
        [System.Windows.Controls.Grid]::SetRow($btnSp, 1)
        [void]$gridMain.Children.Add($btnSp)

        $setWin.Content = $gridMain
        [void]$setWin.ShowDialog()
    })

        # Restore from Backup Button
    $btnRestoreBackup.add_Click({
        $basePath = $txtBasePath.Text
        if ([string]::IsNullOrEmpty($basePath) -or -not (Test-Path $basePath)) {
            [System.Windows.Forms.MessageBox]::Show((Get-UiString 'RestoreNoBackups'), (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $backupDir = $script:AppConfig.BackupDirectory
        if (-not (Test-Path $backupDir)) {
            [System.Windows.Forms.MessageBox]::Show((Get-UiString 'RestoreNoBackups'), (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $baseBaseName = [System.IO.Path]::GetFileNameWithoutExtension($basePath)
        $backups = @(Get-ChildItem -Path $backupDir -Filter "$baseBaseName.*.bak.xlsx" | Sort-Object LastWriteTime -Descending)
        if ($backups.Count -eq 0) {
            $backups = @(Get-ChildItem -Path $backupDir -Filter "*.bak.xlsx" | Sort-Object LastWriteTime -Descending)
        }
        if ($backups.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show((Get-UiString 'RestoreNoBackups'), (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $latest = $backups[0]
        $msg = (Get-UiString 'RestoreConfirmMsg') -f $latest.FullName
        $res = [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'RestoreConfirmTitle'), [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($res -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                Copy-Item -Path $latest.FullName -Destination $basePath -Force
                & $LoadBaseFile $basePath
                [System.Windows.Forms.MessageBox]::Show((Get-UiString 'RestoreSuccess'), (Get-UiString 'InfoTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                                $msg = (Get-UiString 'ErrRestoreBackupMsg') -f $_.Exception.Message
                [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'ErrorTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })

    # Open Backups Folder Action
    if ($btnOpenBackups) {
        $btnOpenBackups.add_Click({
            $bDir = $script:AppConfig.BackupDirectory
            if ([string]::IsNullOrWhiteSpace($bDir)) {
                $baseDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
                $bDir = Join-Path $baseDir 'Backups'
            }
            if (-not (Test-Path $bDir)) { [void][System.IO.Directory]::CreateDirectory($bDir) }
            [System.Diagnostics.Process]::Start('explorer.exe', $bDir) | Out-Null
        })
    }

    # Open Logs Folder Action
    if ($btnOpenLogs) {
        $btnOpenLogs.add_Click({
            $lDir = $script:AppConfig.LogDirectory
            if ([string]::IsNullOrWhiteSpace($lDir)) {
                $baseDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
                $lDir = Join-Path $baseDir 'Logs'
            }
            if (-not (Test-Path $lDir)) { [void][System.IO.Directory]::CreateDirectory($lDir) }
            [System.Diagnostics.Process]::Start('explorer.exe', $lDir) | Out-Null
        })
    }

    # In-App Help Dialog
    $ShowHelpDialog = {
        $p = Get-UpdaterThemePalette $script:CurrentTheme
        $helpWin = New-Object System.Windows.Window -Property @{
            Title                 = Get-UiString 'HelpTitle'
            Width                 = 840
            Height                = 640
            WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner
            Owner                 = $window
            Background            = $script:BrushConverter.ConvertFromString($p.BgApp)
            Foreground            = $script:BrushConverter.ConvertFromString($p.TextPrimary)
            FontFamily            = New-Object System.Windows.Media.FontFamily('Segoe UI')
            FontSize              = 13
        }
        $hHwnd = (New-Object System.Windows.Interop.WindowInteropHelper($helpWin)).EnsureHandle()
        Set-WindowDwmTheme -Hwnd $hHwnd -IsDark $p.IsDark
        if ($null -ne $window.Resources) {
            foreach ($k in $window.Resources.Keys) { if ($null -ne $window.Resources[$k]) { $helpWin.Resources[$k] = $window.Resources[$k] } }
        }

        $grid = New-Object System.Windows.Controls.Grid
        $r0 = New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto }
        $r1 = New-Object System.Windows.Controls.RowDefinition -Property @{ Height = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star) }
        $r2 = New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto }
        [void]$grid.RowDefinitions.Add($r0); [void]$grid.RowDefinitions.Add($r1); [void]$grid.RowDefinitions.Add($r2)

        # Header banner
        $hdr = New-Object System.Windows.Controls.Border -Property @{
            Background      = $script:BrushConverter.ConvertFromString($p.BgHeader)
            Padding         = New-Object System.Windows.Thickness(18, 14, 18, 14)
            BorderBrush     = $script:BrushConverter.ConvertFromString($p.BorderCard)
            BorderThickness = [System.Windows.Thickness]::new(0, 0, 0, 1)
        }
        $spHdr = New-Object System.Windows.Controls.StackPanel
        $txtHdr1 = New-Object System.Windows.Controls.TextBlock -Property @{
            Text        = Get-UiString 'HelpTitle'
            FontSize    = 17
            FontWeight  = [System.Windows.FontWeights]::Bold
            Foreground  = $script:BrushConverter.ConvertFromString($p.TextPrimary)
        }
        $txtHdr2 = New-Object System.Windows.Controls.TextBlock -Property @{
            Text        = Get-UiString 'HelpSubtitle'
            FontSize    = 12
            Foreground  = $script:BrushConverter.ConvertFromString($p.TextSecondary)
            Margin      = New-Object System.Windows.Thickness(0, 3, 0, 0)
        }
        [void]$spHdr.Children.Add($txtHdr1); [void]$spHdr.Children.Add($txtHdr2)
        $hdr.Child = $spHdr
        [System.Windows.Controls.Grid]::SetRow($hdr, 0)
        [void]$grid.Children.Add($hdr)

        # Tab Control
        $tc = New-Object System.Windows.Controls.TabControl -Property @{
            Background      = $script:BrushConverter.ConvertFromString($p.BgCard)
            BorderBrush     = $script:BrushConverter.ConvertFromString($p.BorderCard)
            BorderThickness = [System.Windows.Thickness]::new(0, 1, 0, 1)
        }

        # Helper to create styled cards inside scroll viewer
        $MakeCard = {
            param([string]$Title, [string]$BodyText)
            $b = New-Object System.Windows.Controls.Border -Property @{
                Background      = $script:BrushConverter.ConvertFromString($p.BgCard)
                BorderBrush     = $script:BrushConverter.ConvertFromString($p.BorderCard)
                BorderThickness = [System.Windows.Thickness]::new(1)
                CornerRadius    = [System.Windows.CornerRadius]::new(6)
                Padding         = New-Object System.Windows.Thickness(14, 12, 14, 12)
                Margin          = New-Object System.Windows.Thickness(0, 0, 0, 12)
            }
            $sp = New-Object System.Windows.Controls.StackPanel
            $tbTitle = New-Object System.Windows.Controls.TextBlock -Property @{
                Text       = $Title
                FontWeight = [System.Windows.FontWeights]::SemiBold
                FontSize   = 14
                Foreground = $script:BrushConverter.ConvertFromString($p.AccentBlue)
                Margin     = New-Object System.Windows.Thickness(0, 0, 0, 6)
            }
            $tbBody = New-Object System.Windows.Controls.TextBlock -Property @{
                Text         = $BodyText
                Foreground   = $script:BrushConverter.ConvertFromString($p.TextPrimary)
                TextWrapping = [System.Windows.TextWrapping]::Wrap
                LineHeight   = 18
            }
            [void]$sp.Children.Add($tbTitle)
            [void]$sp.Children.Add($tbBody)
            $b.Child = $sp
            return $b
        }

        # Tab 1: Workflow
        $ti1 = New-Object System.Windows.Controls.TabItem -Property @{ Header = Get-UiString 'HelpTabWorkflow' }
        $sv1 = New-Object System.Windows.Controls.ScrollViewer -Property @{ VerticalScrollBarVisibility = 'Auto'; Margin = New-Object System.Windows.Thickness(16) }
        $sp1 = New-Object System.Windows.Controls.StackPanel

        $t1c1 = & $MakeCard (Get-UiString 'HelpWorkflowCard1Title') (Get-UiString 'HelpWorkflowCard1Body')
        $t1c2 = & $MakeCard (Get-UiString 'HelpWorkflowCard2Title') (Get-UiString 'HelpWorkflowCard2Body')
        $t1c3 = & $MakeCard (Get-UiString 'HelpWorkflowCard3Title') (Get-UiString 'HelpWorkflowCard3Body')
        [void]$sp1.Children.Add($t1c1); [void]$sp1.Children.Add($t1c2); [void]$sp1.Children.Add($t1c3)
        $sv1.Content = $sp1
        $ti1.Content = $sv1
        [void]$tc.Items.Add($ti1)

        # Tab 2: Merge & Compare Modes
        $tiMerge = New-Object System.Windows.Controls.TabItem -Property @{ Header = Get-UiString 'HelpTabMergeModes' }
        $svMerge = New-Object System.Windows.Controls.ScrollViewer -Property @{ VerticalScrollBarVisibility = 'Auto'; Margin = New-Object System.Windows.Thickness(16) }
        $spMerge = New-Object System.Windows.Controls.StackPanel

        $tmc1 = & $MakeCard (Get-UiString 'HelpMergeModesCardTitle') (Get-UiString 'HelpMergeModesCardBody')
        $tmc2 = & $MakeCard (Get-UiString 'HelpCompareOptionsCardTitle') (Get-UiString 'HelpCompareOptionsCardBody')
        $tmc3 = & $MakeCard (Get-UiString 'HelpSmartAutoMapCardTitle') (Get-UiString 'HelpSmartAutoMapCardBody')
        [void]$spMerge.Children.Add($tmc1); [void]$spMerge.Children.Add($tmc2); [void]$spMerge.Children.Add($tmc3)
        $svMerge.Content = $spMerge
        $tiMerge.Content = $svMerge
        [void]$tc.Items.Add($tiMerge)

        # Tab 3: Backups & Logs
        $ti2 = New-Object System.Windows.Controls.TabItem -Property @{ Header = Get-UiString 'HelpTabBackupsLogs' }
        $sv2 = New-Object System.Windows.Controls.ScrollViewer -Property @{ VerticalScrollBarVisibility = 'Auto'; Margin = New-Object System.Windows.Thickness(16) }
        $sp2 = New-Object System.Windows.Controls.StackPanel

        $curBackups = $script:AppConfig.BackupDirectory
        if ([string]::IsNullOrWhiteSpace($curBackups)) {
            $baseDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
            $curBackups = Join-Path $baseDir 'Backups'
        }
        $curLogs    = $script:AppConfig.LogDirectory
        if ([string]::IsNullOrWhiteSpace($curLogs)) {
            $baseDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
            $curLogs = Join-Path $baseDir 'Logs'
        }

        $t2c1 = & $MakeCard (Get-UiString 'HelpBackupsCardTitle') ((Get-UiString 'HelpBackupsCardBody') -f $curBackups)
        $t2c2 = & $MakeCard (Get-UiString 'HelpLogsCardTitle') ((Get-UiString 'HelpLogsCardBody') -f $curLogs)
        [void]$sp2.Children.Add($t2c1); [void]$sp2.Children.Add($t2c2)
        $sv2.Content = $sp2
        $ti2.Content = $sv2
        [void]$tc.Items.Add($ti2)

        # Tab 3: Shortcuts
        $ti3 = New-Object System.Windows.Controls.TabItem -Property @{ Header = Get-UiString 'HelpTabShortcuts' }
        $sv3 = New-Object System.Windows.Controls.ScrollViewer -Property @{ VerticalScrollBarVisibility = 'Auto'; Margin = New-Object System.Windows.Thickness(16) }
        $sp3 = New-Object System.Windows.Controls.StackPanel

        $t3c1 = & $MakeCard (Get-UiString 'HelpShortcutsCardTitle') (Get-UiString 'HelpShortcutsCardBody')
        [void]$sp3.Children.Add($t3c1)
        $sv3.Content = $sp3
        $ti3.Content = $sv3
        [void]$tc.Items.Add($ti3)

        [System.Windows.Controls.Grid]::SetRow($tc, 1)
        [void]$grid.Children.Add($tc)

        # Bottom Action Bar
        $botBar = New-Object System.Windows.Controls.Border -Property @{
            Background      = $script:BrushConverter.ConvertFromString($p.BgHeader)
            Padding         = New-Object System.Windows.Thickness(16, 10, 16, 10)
            BorderBrush     = $script:BrushConverter.ConvertFromString($p.BorderCard)
            BorderThickness = [System.Windows.Thickness]::new(0, 1, 0, 0)
        }
        $spBot = New-Object System.Windows.Controls.Grid
        $cLeft = New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star) }
        $cRight = New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::Auto }
        [void]$spBot.ColumnDefinitions.Add($cLeft); [void]$spBot.ColumnDefinitions.Add($cRight)

        $spBotLeft = New-Object System.Windows.Controls.StackPanel -Property @{ Orientation = [System.Windows.Controls.Orientation]::Horizontal }
        
        $btnDoc = New-Object System.Windows.Controls.Button -Property @{
            Content         = Get-UiString 'HelpBtnOpenDoc'
            Background      = $script:BrushConverter.ConvertFromString($p.BtnSecondaryBg)
            Foreground      = $script:BrushConverter.ConvertFromString($p.BtnSecondaryFg)
            BorderBrush     = $script:BrushConverter.ConvertFromString($p.BorderCard)
            BorderThickness = [System.Windows.Thickness]::new(1)
            Padding         = New-Object System.Windows.Thickness(12, 6, 12, 6)
            Margin          = New-Object System.Windows.Thickness(0, 0, 8, 0)
        }
        $btnDoc.add_Click({
            $baseDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
            $langSuffix = switch ($script:CurrentLanguage) {
                'pl' { '_PL.md' }
                'de' { '_DE.md' }
                default { '_EN.md' }
            }
            $targetDoc = Join-Path $baseDir "Docs\USER_GUIDE$langSuffix"
            $docPath = if (Test-Path $targetDoc) { $targetDoc } else { Join-Path $baseDir 'Docs\USER_GUIDE.md' }
            if (Test-Path $docPath) {
                try {
                    $psi = New-Object System.Diagnostics.ProcessStartInfo $docPath
                    $psi.UseShellExecute = $true
                    [System.Diagnostics.Process]::Start($psi) | Out-Null
                } catch {
                    try {
                        Start-Process 'notepad.exe' -ArgumentList "\"$docPath\""
                    } catch {
                        $errMsg = $_.Exception.Message
                        [System.Windows.Forms.MessageBox]::Show($errMsg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    }
                }
            } else {
                                $msg = (Get-UiString 'ErrUserGuideNotFound') -f $docPath
                [System.Windows.Forms.MessageBox]::Show($msg, (Get-UiString 'WarningTitle'), [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
        })

        $btnHBackups = New-Object System.Windows.Controls.Button -Property @{
            Content         = Get-UiString 'BtnOpenBackups'
            Background      = $script:BrushConverter.ConvertFromString($p.BtnSecondaryBg)
            Foreground      = $script:BrushConverter.ConvertFromString($p.BtnSecondaryFg)
            BorderBrush     = $script:BrushConverter.ConvertFromString($p.BorderCard)
            BorderThickness = [System.Windows.Thickness]::new(1)
            Padding         = New-Object System.Windows.Thickness(12, 6, 12, 6)
            Margin          = New-Object System.Windows.Thickness(0, 0, 8, 0)
        }
        $btnHBackups.add_Click({
            $bDir = $script:AppConfig.BackupDirectory
            if ([string]::IsNullOrWhiteSpace($bDir)) {
                $baseDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
                $bDir = Join-Path $baseDir 'Backups'
            }
            if (-not (Test-Path $bDir)) { [void][System.IO.Directory]::CreateDirectory($bDir) }
            [System.Diagnostics.Process]::Start('explorer.exe', $bDir) | Out-Null
        })

        $btnHLogs = New-Object System.Windows.Controls.Button -Property @{
            Content         = Get-UiString 'BtnOpenLogs'
            Background      = $script:BrushConverter.ConvertFromString($p.BtnSecondaryBg)
            Foreground      = $script:BrushConverter.ConvertFromString($p.BtnSecondaryFg)
            BorderBrush     = $script:BrushConverter.ConvertFromString($p.BorderCard)
            BorderThickness = [System.Windows.Thickness]::new(1)
            Padding         = New-Object System.Windows.Thickness(12, 6, 12, 6)
        }
        $btnHLogs.add_Click({
            $lDir = $script:AppConfig.LogDirectory
            if ([string]::IsNullOrWhiteSpace($lDir)) {
                $baseDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
                $lDir = Join-Path $baseDir 'Logs'
            }
            if (-not (Test-Path $lDir)) { [void][System.IO.Directory]::CreateDirectory($lDir) }
            [System.Diagnostics.Process]::Start('explorer.exe', $lDir) | Out-Null
        })

        [void]$spBotLeft.Children.Add($btnDoc)
        [void]$spBotLeft.Children.Add($btnHBackups)
        [void]$spBotLeft.Children.Add($btnHLogs)
        [System.Windows.Controls.Grid]::SetColumn($spBotLeft, 0)
        [void]$spBot.Children.Add($spBotLeft)

        $btnClose = New-Object System.Windows.Controls.Button -Property @{
            Content         = Get-UiString 'HelpBtnClose'
            Background      = $script:BrushConverter.ConvertFromString($p.AccentBlue)
            Foreground      = $script:BrushConverter.ConvertFromString('#FFFFFF')
            FontWeight      = [System.Windows.FontWeights]::SemiBold
            Padding         = New-Object System.Windows.Thickness(18, 6, 18, 6)
        }
        $btnClose.add_Click({ $helpWin.Close() })
        [System.Windows.Controls.Grid]::SetColumn($btnClose, 1)
        [void]$spBot.Children.Add($btnClose)

        $botBar.Child = $spBot
        [System.Windows.Controls.Grid]::SetRow($botBar, 2)
        [void]$grid.Children.Add($botBar)

        $helpWin.Content = $grid
        [void]$helpWin.ShowDialog()
    }

    if ($btnHelp) {
        $btnHelp.add_Click({
            & $ShowHelpDialog
        })
    }

    # Initial Localization & Theme application
    & $UpdateLocalization

    $shouldLoadBase = if (-not [string]::IsNullOrEmpty($InitBaseFilePath)) {
        Test-Path $InitBaseFilePath
    } else {
        ($cfg.RememberBasePath -ne $false) -and (-not [string]::IsNullOrEmpty($cfg.BaseFilePath)) -and (Test-Path $cfg.BaseFilePath)
    }
    if ($shouldLoadBase) {
        $pathToLoad = if (-not [string]::IsNullOrEmpty($InitBaseFilePath)) { $InitBaseFilePath } else { $cfg.BaseFilePath }
        & $LoadBaseFile $pathToLoad
    }
    if (-not [string]::IsNullOrEmpty($InitIncomingPath) -and (Test-Path $InitIncomingPath)) {
        & $LoadIncomingFile $InitIncomingPath
    }

    # Ensure sheet headers and matching profile are loaded if initial files were provided
    if ($cmbBaseSheet.SelectedItem -and (Test-Path $txtBasePath.Text) -and (-not $script:BaseHeaders -or $script:BaseHeaders.Count -eq 0)) {
        try {
            $script:BaseHeaders = [FastExcelHelper]::GetHeaders($txtBasePath.Text, $cmbBaseSheet.SelectedItem.ToString())
            $lbJoinBase.Items.Clear()
            foreach ($h in $script:BaseHeaders) { [void]$lbJoinBase.Items.Add($h) }
        } catch { }
    }
    if ($cmbIncomingSheet.SelectedItem -and (Test-Path $txtIncomingPath.Text) -and (-not $script:IncomingHeaders -or $script:IncomingHeaders.Count -eq 0)) {
        try {
            $script:IncomingHeaders = [FastExcelHelper]::GetHeaders($txtIncomingPath.Text, $cmbIncomingSheet.SelectedItem.ToString())
            $lbJoinIncoming.Items.Clear()
            foreach ($h in $script:IncomingHeaders) { [void]$lbJoinIncoming.Items.Add($h) }
            & $CheckProfileAutoMatch
        } catch { }
    }

    # Show Window if not non-interactive
    if (-not $NonInteractive) {
        try {
            $window.ShowDialog() | Out-Null
        } catch {
            $baseDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
            $crashLog = Join-Path $baseDir 'crash.log'
            $info = "SHOWDIALOG CRASH:`n" +
                    "Exception: $($_.Exception.ToString())`n" +
                    "ScriptStackTrace:`n$($_.ScriptStackTrace)`n" +
                    "Position:`n$($_.InvocationInfo.PositionMessage)`n" +
                    "Inner: $($_.Exception.InnerException)`n"
            [System.IO.File]::WriteAllText($crashLog, $info, [System.Text.UTF8Encoding]::new($true))
            Write-Host $info -ForegroundColor Red
            throw
        }
    } else {
        return $window
    }
}

# Auto-launch if executed directly
if ($MyInvocation.InvocationName -ne '.' -and (-not $MyInvocation.Line -or -not $MyInvocation.Line.StartsWith('.'))) {
    if ($Headless -or ($AutoAccept -ne 'None' -and -not [string]::IsNullOrEmpty($BaseFilePath) -and -not [string]::IsNullOrEmpty($IncomingPath))) {
        if ('ConsoleHelper' -as [type]) {
            try { [void][ConsoleHelper]::AttachConsole(-1) } catch { }
        }
        $hRes = Invoke-HeadlessMasterUpdater -BaseFilePath $BaseFilePath -IncomingPath $IncomingPath -BaseSheet $BaseSheet -IncomingSheet $IncomingSheet -ConfigPath $ConfigPath -AutoAccept $AutoAccept -ExportReportPath $ExportReportPath -SummaryJsonPath $SummaryJsonPath -Quiet:$Quiet
        if ($MyInvocation.MyCommand.Path -match '\.exe$' -or [System.AppDomain]::CurrentDomain.FriendlyName -match '\.exe$') {
            $ec = if ($hRes -and $hRes.Success) { 0 } else { 1 }
            exit $ec
        }
    } else {
        Show-MasterUpdater -InitBaseFilePath $BaseFilePath -InitIncomingPath $IncomingPath -InitBaseSheet $BaseSheet -InitIncomingSheet $IncomingSheet -CustomConfigPath $ConfigPath -NonInteractive:$NonInteractive
    }
}
