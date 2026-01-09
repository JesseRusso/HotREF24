using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace HotPort.Models
{
    public static class ExcelHelper
    {
        private static readonly Dictionary<string, (FileStream Stream, SpreadsheetDocument Document)> cache = new();

        public static string GetCellValue(string filePath, string sheetName, string cellReference)
        {
            var doc = GetDocument(filePath);
            WorkbookPart? wbPart = doc.WorkbookPart;
            Sheet? theSheet = wbPart?.Workbook.Descendants<Sheet>()
                .FirstOrDefault(s => s.Name == sheetName);
            if (theSheet == null)
            {
                throw new ArgumentException(null, nameof(sheetName));
            }
            WorksheetPart wsPart = (WorksheetPart)wbPart!.GetPartById(theSheet.Id!);
            Cell? theCell = wsPart.Worksheet.Descendants<Cell>()
                .FirstOrDefault(c => c.CellReference == cellReference);
            if (theCell == null || string.IsNullOrEmpty(theCell.InnerText))
            {
                return string.Empty;
            }
            string? value = theCell.CellValue?.InnerText;
            if (theCell.DataType != null)
            {
                if (theCell.DataType.Value == CellValues.SharedString)
                {
                    var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    if (stringTable != null)
                    {
                        value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                    }
                }
                else if (theCell.DataType.Value == CellValues.Boolean)
                {
                    value = value == "0" ? "FALSE" : "TRUE";
                }
            }
            return value ?? string.Empty;
        }

        private static SpreadsheetDocument GetDocument(string filePath)
        {
            if (!cache.TryGetValue(filePath, out var entry))
            {
                var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                var document = SpreadsheetDocument.Open(stream, false);
                entry = (stream, document);
                cache[filePath] = entry;
            }
            return entry.Document;
        }

        public static void CloseCachedDocuments()
        {
            foreach (var entry in cache.Values)
            {
                entry.Document.Dispose();
                entry.Stream.Close();
            }
            cache.Clear();
        }
    }
}