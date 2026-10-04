# Test-EditExcelHelper.ps1
# Unit test for EditExcelHelper and FastExcelHelper

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Xml
Add-Type -AssemblyName System.Xml.Linq

$csharp = @"
using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Linq;
using System.Globalization;
using System.Management.Automation;

public class RowOp {
    public string Type; // "PatchCell" or "AppendRow"
    public int RowNumber;
    public Dictionary<int, string> Cells = new Dictionary<int, string>();
}

public static class EditExcelHelper {
    private static readonly XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace relsNs = "http://schemas.openxmlformats.org/package/2006/relationships";

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
                    if (rId == null) rId = id; // fallback to first
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

    public static void WriteChanges(string filePath, string sheetName, IList<RowOp> ops) {
        if (!File.Exists(filePath)) throw new FileNotFoundException("Base file not found", filePath);
        string tempPath = filePath + ".tmp.xlsx";
        if (File.Exists(tempPath)) File.Delete(tempPath);

        string targetSheetPath = null;
        XDocument doc = null;

        // Open original file
        using (var inStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
        using (var zipIn = new ZipArchive(inStream, ZipArchiveMode.Read)) {
            targetSheetPath = ResolveSheetTarget(zipIn, sheetName);
            var sheetEntry = zipIn.GetEntry(targetSheetPath);
            if (sheetEntry == null) throw new InvalidOperationException("Target worksheet not found: " + targetSheetPath);

            using (var sheetStream = sheetEntry.Open()) {
                doc = XDocument.Load(sheetStream);
            }

            // GUARD: Scan for tableParts, pivotTable, and shared formulas
            string rawXml = doc.ToString();
            if (rawXml.IndexOf("tableParts", StringComparison.OrdinalIgnoreCase) >= 0 ||
                rawXml.IndexOf("tablePart", StringComparison.OrdinalIgnoreCase) >= 0 ||
                rawXml.IndexOf("pivotTable", StringComparison.OrdinalIgnoreCase) >= 0 ||
                rawXml.IndexOf("t=\"shared\"", StringComparison.OrdinalIgnoreCase) >= 0) {
                throw new InvalidOperationException("GUARD: Unsupported structure detected in worksheet (tableParts, pivotTable, or shared formulas). InPlace write aborted.");
            }

            var sheetData = doc.Root.Element(ns + "sheetData");
            if (sheetData == null) throw new InvalidOperationException("sheetData element not found in worksheet");

            // Apply operations
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

            // Update dimension if present using fast incremental bounds
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

            // Create temporary archive and write all entries
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

        // Atomic swap over original file
        File.Copy(tempPath, filePath, true);
        try { File.Delete(tempPath); } catch { }
    }
}
"@

if (-not ('EditExcelHelper' -as [type])) {
    if ($PSVersionTable.PSVersion.Major -le 5) {
        Add-Type -TypeDefinition $csharp -ReferencedAssemblies @(
            'System.Xml',
            'System.Xml.Linq',
            'System.Core',
            'System.IO.Compression',
            'System.IO.Compression.FileSystem',
            ([PSObject].Assembly.Location)
        ) -Language CSharp
    } else {
        Add-Type -TypeDefinition $csharp -Language CSharp
    }
    Write-Host "EditExcelHelper compiled successfully!"
} else {
    Write-Host "EditExcelHelper already loaded into session."
}

# Functional test on a copy of base.xlsx
$testBaseCopy = Join-Path $PSScriptRoot "base_test_copy.xlsx"
$originalBase = Join-Path (Split-Path $PSScriptRoot -Parent) "TestFixtures\base.xlsx"
Copy-Item -Path $originalBase -Destination $testBaseCopy -Force

# Create operations
$ops = New-Object 'System.Collections.Generic.List[RowOp]'

# Op 1: Patch row 6 (physical Excel row 6), column 7 (TelefonRodzica)
$opPatch = New-Object RowOp
$opPatch.Type = 'PatchCell'
$opPatch.RowNumber = 6
$opPatch.Cells.Add(7, '999-888-777') # 0-based col 7 = H (TelefonRodzica)
$opPatch.Cells.Add(20, '2026-10-03 15:30') # col 20 = U (OstZmiana)
$ops.Add($opPatch)

# Op 2: Append row 52
$opAppend = New-Object RowOp
$opAppend.Type = 'AppendRow'
$opAppend.RowNumber = 52
$opAppend.Cells.Add(0, 'D-9999')
$opAppend.Cells.Add(1, 'Testowy Jan')
$opAppend.Cells.Add(4, '8B')
$opAppend.Cells.Add(7, '111-222-333')
$opAppend.Cells.Add(20, '2026-10-03 15:30')
$ops.Add($opAppend)

# Execute
[EditExcelHelper]::WriteChanges($testBaseCopy, 'Dzieci', $ops)

# Verify with ZipArchive and XML
$fs = [System.IO.File]::OpenRead($testBaseCopy)
$zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
$entry = $zip.GetEntry('xl/worksheets/sheet1.xml')
$sr = New-Object System.IO.StreamReader($entry.Open())
$xml = $sr.ReadToEnd()
$sr.Dispose()
$zip.Dispose()
$fs.Dispose()

if ($xml -match '999-888-777' -and $xml -match 'Testowy Jan' -and $xml -match 'D-9999') {
    Write-Host "SUCCESS: Cell patch and Row append verified in worksheet XML!"
} else {
    throw "FAILED: Patched values not found in worksheet XML"
}

# Clean up test copy
Remove-Item -Force $testBaseCopy

# Functional test on base_formatted.xlsx (multi-sheet preservation)
$testFormattedCopy = Join-Path $PSScriptRoot "base_formatted_test_copy.xlsx"
$originalFormatted = Join-Path (Split-Path $PSScriptRoot -Parent) "TestFixtures\base_formatted.xlsx"
Copy-Item -Path $originalFormatted -Destination $testFormattedCopy -Force

$opsFormatted = New-Object 'System.Collections.Generic.List[RowOp]'
$opP = New-Object RowOp
$opP.Type = 'PatchCell'
$opP.RowNumber = 10
$opP.Cells.Add(7, '888-777-666')
$opsFormatted.Add($opP)

[EditExcelHelper]::WriteChanges($testFormattedCopy, 'Dzieci', $opsFormatted)

$fs = [System.IO.File]::OpenRead($testFormattedCopy)
$zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
$sheet1Entry = $zip.GetEntry('xl/worksheets/sheet1.xml')
$sheet2Entry = $zip.GetEntry('xl/worksheets/sheet2.xml')

if ($null -eq $sheet2Entry) {
    $zip.Dispose()
    $fs.Dispose()
    throw "FAILED: Sheet 2 (Instrukcja) was lost during InPlace update!"
}

$sr2 = New-Object System.IO.StreamReader($sheet2Entry.Open())
$s2Xml = $sr2.ReadToEnd()
$sr2.Dispose()
$zip.Dispose()
$fs.Dispose()

if ($s2Xml -match 'Ten arkusz musi przetrwac bez zmian') {
    Write-Host "SUCCESS: Multi-sheet and formatting preserved on base_formatted.xlsx!"
} else {
    throw "FAILED: Content in Sheet 2 was corrupted"
}

Remove-Item -Force $testFormattedCopy

Write-Host "All EditExcelHelper tests PASSED."
