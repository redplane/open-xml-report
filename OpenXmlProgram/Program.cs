using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlProgram
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var output = Path.Combine(path, $"{DateTime.Now.ToString("yy-MM-dd HH-mm-ss")}.xlsx");
            path = Path.Combine(path, "Files/bkt-14.xlsx");

            var input = File.ReadAllBytes(path);
            var memoryStream = new MemoryStream(input);

            using (var spreadSheetDocument = SpreadsheetDocument.Open(memoryStream, true))
            {
                var sheets = FindSheets(spreadSheetDocument, x => x.Name.Value.Contains("Trang"));
                foreach (var sheet in sheets)
                {
                    var part = FindWorksheetPart(spreadSheetDocument, sheet);
                    var rows = FindRows(part.Worksheet);
                    #region Station name

                    var cellStationNames = FindCells(rows).Where(c => string.Compare
                                                    (c.CellReference.Value, "E1", StringComparison.InvariantCultureIgnoreCase) == 0).ToList();

                    cellStationNames.ForEach(
                        x =>
                        {
                            x.CellValue = new CellValue("This is not station name");
                            x.DataType = new EnumValue<CellValues>(CellValues.String); 
                        });

                    #endregion
                    

                    part.Worksheet.Save();
                }
                
                spreadSheetDocument.SaveAs(output);
            }
        }

        /// <summary>
        /// Find list of sheets by using specific conditions.
        /// </summary>
        /// <param name="spreadsheetDocument"></param>
        /// <param name="condition"></param>
        /// <returns></returns>
        private static IList<Sheet> FindSheets(SpreadsheetDocument spreadsheetDocument, Func<Sheet, bool> condition)
        {
            // Find list of sheets.
            return
                spreadsheetDocument.WorkbookPart.Workbook.Sheets.Elements<Sheet>().Where(condition).ToList();
        }

        /// <summary>
        /// Find worksheet part by using sheet information.
        /// </summary>
        /// <param name="spreadsheetDocument"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private static WorksheetPart FindWorksheetPart(SpreadsheetDocument spreadsheetDocument, Sheet sheet)
        {
            var worksheetPart = (WorksheetPart)
                 spreadsheetDocument.WorkbookPart.GetPartById(sheet.Id.Value);
            return worksheetPart;
        }

        /// <summary>
        /// Find rows in worksheet by using specific conditions.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        private static IEnumerable<Row> FindRows(Worksheet worksheet)
        {
            return worksheet.Elements<SheetData>().SelectMany(x => x.Elements<Row>());
        }

        /// <summary>
        /// Find rows in worksheet by using specific conditions.
        /// </summary>
        /// <param name="worksheets"></param>
        /// <param name="condition"></param>
        /// <returns></returns>
        private static IEnumerable<Row> FindRows(IEnumerable<Worksheet> worksheets, Func<Row, bool> condition)
        {
            return worksheets.SelectMany(worksheet => worksheet.Elements<SheetData>().SelectMany(x => x.Elements<Row>().Where(condition)));
        }

        /// <summary>
        /// Find cells from collection of rows.
        /// </summary>
        /// <param name="rows"></param>
        /// <returns></returns>
        private static IEnumerable<Cell> FindCells(IEnumerable<Row> rows)
        {
            return rows.SelectMany(x => x.Elements<Cell>());
        }


        /// <summary>
        /// Find cells by using specific conditions from collection of rows.
        /// </summary>
        /// <param name="rows"></param>
        /// <param name="condition"></param>
        /// <returns></returns>
        private static IEnumerable<Cell> FindCells(IEnumerable<Row> rows , Func<Cell, bool> condition)
        {
            return rows.SelectMany(x => x.Elements<Cell>().Where(condition));
        }
        
    }
}