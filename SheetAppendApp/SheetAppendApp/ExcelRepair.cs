using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SheetAppendApp
{
    public static class ExcelRepair
    {
        public static void RemoveBrokenReferences(string xlsxPath, Action<string>? log = null)
        {
            log ??= _ => { };

            using var doc = SpreadsheetDocument.Open(xlsxPath, true);
            var wbPart = doc.WorkbookPart!;
            var wb = wbPart.Workbook;

            var existingSheets = wb.Sheets!.Elements<Sheet>()
                .Select(s => s.Name?.Value ?? "")
                .Where(n => n.Length > 0)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            int removedDefinedNames = RemoveBrokenDefinedNames(wb, existingSheets);
            int removedDV = 0;
            int removedCF = 0;

            foreach (var wsPart in wbPart.WorksheetParts)
            {
                var ws = wsPart.Worksheet;

                removedDV += RemoveBrokenDataValidations(ws, existingSheets);
                removedCF += RemoveBrokenConditionalFormatting(ws, existingSheets);

                ws.Save();
            }

            wb.Save();

            if (removedDefinedNames > 0) log($"[REPAIR] Removed DefinedNames: {removedDefinedNames}");
            if (removedDV > 0) log($"[REPAIR] Removed DataValidations: {removedDV}");
            if (removedCF > 0) log($"[REPAIR] Removed ConditionalFormatting rules: {removedCF}");

            if (removedDefinedNames == 0 && removedDV == 0 && removedCF == 0)
                log("[REPAIR] No broken references found (or nothing removable).");
        }

        private static int RemoveBrokenDefinedNames(Workbook wb, HashSet<string> existingSheets)
        {
            var definedNames = wb.DefinedNames;
            if (definedNames == null) return 0;

            var list = definedNames.Elements<DefinedName>().ToList();
            int removed = 0;

            foreach (var dn in list)
            {
                var text = dn.Text ?? "";
                if (IsBrokenRef(text, existingSheets))
                {
                    dn.Remove();
                    removed++;
                }
            }

            if (!definedNames.Elements<DefinedName>().Any())
                definedNames.Remove();

            return removed;
        }

        private static int RemoveBrokenDataValidations(Worksheet ws, HashSet<string> existingSheets)
        {
            var dvs = ws.Elements<DataValidations>().FirstOrDefault();
            if (dvs == null) return 0;

            var dvList = dvs.Elements<DataValidation>().ToList();
            int removed = 0;

            foreach (var dv in dvList)
            {
                // ✅ FIX: ListValue<> không có .Value -> dùng InnerText
                var seq = dv.SequenceOfReferences?.InnerText;

                bool bad =
                    IsBrokenRef(dv.Formula1?.Text, existingSheets) ||
                    IsBrokenRef(dv.Formula2?.Text, existingSheets) ||
                    IsBrokenRef(seq, existingSheets);

                if (bad)
                {
                    dv.Remove();
                    removed++;
                }
            }

            if (!dvs.Elements<DataValidation>().Any())
                dvs.Remove();

            return removed;
        }

        private static int RemoveBrokenConditionalFormatting(Worksheet ws, HashSet<string> existingSheets)
        {
            // ✅ FIX: không dùng CfRule/CFRule (tùy version có/không)
            var cfs = ws.Elements<ConditionalFormatting>().ToList();
            if (cfs.Count == 0) return 0;

            int removedRules = 0;

            foreach (var cf in cfs)
            {
                // cfRule nằm dưới ConditionalFormatting, nhưng kiểu có thể khác theo version
                var rules = cf.ChildElements
                    .Where(e => string.Equals(e.LocalName, "cfRule", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                foreach (var rule in rules)
                {
                    bool bad = rule.Descendants<Formula>()
                        .Any(f => IsBrokenRef(f.Text, existingSheets));

                    if (bad)
                    {
                        rule.Remove();
                        removedRules++;
                    }
                }

                bool hasAnyRule = cf.ChildElements.Any(e => string.Equals(e.LocalName, "cfRule", StringComparison.OrdinalIgnoreCase));
                if (!hasAnyRule)
                    cf.Remove();
            }

            return removedRules;
        }

        private static bool IsBrokenRef(string? text, HashSet<string> existingSheets)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;

            if (text.Contains("#REF!", StringComparison.OrdinalIgnoreCase))
                return true;

            foreach (var sheet in ExtractSheetRefs(text))
            {
                if (!existingSheets.Contains(sheet))
                    return true;
            }
            return false;
        }

        private static IEnumerable<string> ExtractSheetRefs(string text)
        {
            // 'Sheet Name'!A1  OR  SheetName!A1  (sheet name có thể unicode)
            var rx = new Regex(@"(?:'([^']+)'|([^!]+))!", RegexOptions.Compiled);
            foreach (Match m in rx.Matches(text))
            {
                var name = m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value;
                name = name.Trim();
                if (!string.IsNullOrWhiteSpace(name))
                    yield return name;
            }
        }
    }
}