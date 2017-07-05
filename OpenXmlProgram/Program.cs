using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Xml.XPath;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using OpenXmlProgram.ViewModels.Report.BktFourteen;

namespace OpenXmlProgram
{
    internal class Program
    {

        private static void Main(string[] args)
        {
            // Find application directory.
            var path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            // Path to which file should be exported.
            var output = Path.Combine(path, $"{DateTime.Now.ToString("yy-MM-dd HH-mm-ss")}.xlsx");

            // Template file.
            path = Path.Combine(path, "Files/bkt-14.xlsx");

            // Read all template data.
            var input = File.ReadAllBytes(path);
            var memoryStream = new MemoryStream();
            memoryStream.Write(input, 0, input.Length);

            // TODO: Remove
            var bktReport = FindReport();

            if (bktReport.Station == null)
            {
                Console.WriteLine("Invalid station");
                Console.ReadLine();
                return;
            }

            // Find station in report.
            var station = bktReport.Station;

            // Data source sheet.
            const string sheetDataSource = "Binding";

            using (var spreadSheetDocument = SpreadsheetDocument.Open(memoryStream, true))
            {
                // Find binding sheet.
                var sheet = FindSheets(spreadSheetDocument, x => x.Name.Value.Equals(sheetDataSource)).FirstOrDefault();
                spreadSheetDocument.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
                spreadSheetDocument.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;

                // Find sheet attached information.
                var part = FindWorksheetPart(spreadSheetDocument, sheet);
                var rows = FindRows(part.Worksheet);

                #region General information

                var cellStationName = FindCell(rows, "E", 1);
                if (cellStationName != null)
                {
                    cellStationName.CellValue = new CellValue(station.Name);
                    cellStationName.DataType = new EnumValue<CellValues>(CellValues.String);
                }

                var cellMonth = FindCell(rows, "K", 1);
                if (cellMonth != null)
                {
                    cellMonth.CellValue = new CellValue(bktReport.Time.Month.ToString());
                    cellMonth.DataType = new EnumValue<CellValues>(CellValues.Number);
                }

                var cellYear = FindCell(rows, "K", 1);
                if (cellYear != null)
                {
                    cellYear.CellValue = new CellValue(bktReport.Time.Year.ToString());
                    cellYear.DataType = new EnumValue<CellValues>(CellValues.Number);
                }
                
                #endregion

                #region Report information.

                var time = bktReport.Time;
                uint row = 6;

                foreach (var dayReport in bktReport.Reports)
                {
                    for (var column = 'C'; column < 'Z'; column++)
                    {
                        var iColumnIndex = column - 'C';
                        var cellRain = InsertCell($"{column}", row, part);
                        var cellRainObs = InsertCell($"{column}", row + 1, part);

                        cellRain.CellValue = new CellValue(dayReport[iColumnIndex].Rain.ToString());
                        cellRain.DataType = new EnumValue<CellValues>(CellValues.Number);
                        cellRainObs.CellValue = new CellValue(dayReport[iColumnIndex].RainObs.ToString());
                        cellRainObs.DataType = new EnumValue<CellValues>(CellValues.Number);
                    }

                    row += 2;
                }

                #endregion

                part.Worksheet.Save();
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
        private static IEnumerable<Cell> FindCells(IEnumerable<Row> rows, Func<Cell, bool> condition)
        {
            return rows.SelectMany(x => x.Elements<Cell>().Where(condition));
        }

        /// <summary>
        /// Find report information.
        /// </summary>
        /// <returns></returns>
        private static BktFourteenReportViewModel FindReport()
        {
            var path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            path = Path.Combine(path, "Data/bkt-14.json");

            var text = File.ReadAllText(path);
            return JsonConvert.DeserializeObject<BktFourteenReportViewModel>(text);
        }

        /// <summary>
        /// Find cell at a specific position.
        /// </summary>
        /// <param name="rows"></param>
        /// <param name="column"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private static Cell FindCell(IEnumerable<Row> rows, string column, uint row)
        {
            return FindCells(rows).FirstOrDefault(c => string.Compare
                                                                (c.CellReference.Value, $"{column}{row}", StringComparison.InvariantCultureIgnoreCase) == 0);
        }

        private static Cell InsertCell(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            var worksheet = worksheetPart.Worksheet;
            var sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;
            Row row;
            if (sheetData.Elements<Row>().Count(r => r.RowIndex == rowIndex) != 0)
            {
                row = sheetData.Elements<Row>().First(r => r.RowIndex == rowIndex);
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }
            if (row.Elements<Cell>().Any(c => c.CellReference.Value == columnName + rowIndex))
            {
                return row.Elements<Cell>().First(c => c.CellReference.Value == cellReference);
            }
            else
            {
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (String.Compare(cell.CellReference.Value, cellReference, StringComparison.OrdinalIgnoreCase) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }
                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);
                worksheet.Save();
                return newCell;
            }
        }
    }
}