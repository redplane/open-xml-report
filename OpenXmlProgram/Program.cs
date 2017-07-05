using System;
using System.Collections.Generic;
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
            var path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var output = Path.Combine(path, $"{DateTime.Now.ToString("yy-MM-dd HH-mm-ss")}.xlsx");
            var xmlFile = Path.Combine(path, "Data/bkt-14.xml");
            var xml = File.ReadAllText(xmlFile);
            
            using (var spreadSheetDocument = SpreadsheetDocument.Create(output, SpreadsheetDocumentType.Workbook))
            {
                var workbookpart = spreadSheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();
                var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();

                var stream = worksheetPart.GetStream();
                var bytes = Encoding.UTF8.GetBytes(xml);
                stream.Write(bytes, 0, bytes.Length);
                workbookpart.Workbook.Save();
                spreadSheetDocument.Close();
            }
            
        }

        protected static void AddPartXml(OpenXmlPart part, string xml)
        {
            using (Stream stream = part.GetStream())
            {
                byte[] buffer = (new UTF8Encoding()).GetBytes(xml);
                stream.Write(buffer, 0, buffer.Length);
            }
        }

        private static void Main1(string[] args)
        {
            var path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var output = Path.Combine(path, $"{DateTime.Now.ToString("yy-MM-dd HH-mm-ss")}.xlsx");
            path = Path.Combine(path, "Files/bkt-14.xlsx");

            var input = File.ReadAllBytes(path);
            var memoryStream = new MemoryStream(input);

            // TODO: Remove
            var bktReport = FindReport();

            using (var spreadSheetDocument = SpreadsheetDocument.Open(memoryStream, true))
            {
                var sheets = FindSheets(spreadSheetDocument, x => x.Name.Value.Contains("Trang"));
                foreach (var sheet in sheets)
                {
                    var part = FindWorksheetPart(spreadSheetDocument, sheet);
                    var rows = FindRows(part.Worksheet);

                    #region General information

                    var cellStationNames = FindCells(rows).Where(c => string.Compare
                                                    (c.CellReference.Value, "E1", StringComparison.InvariantCultureIgnoreCase) == 0).ToList();

                    var cellMonths = FindCells(rows).Where(c => string.Compare
                                                                  (c.CellReference.Value, "K1", StringComparison.InvariantCultureIgnoreCase) == 0).ToList();

                    var cellYears = FindCells(rows).Where(c => string.Compare
                                                                    (c.CellReference.Value, "O1", StringComparison.InvariantCultureIgnoreCase) == 0).ToList();


                    cellStationNames.ForEach(
                        x =>
                        {
                            x.CellValue = new CellValue(bktReport.Station.Name);
                            x.DataType = new EnumValue<CellValues>(CellValues.String);
                        });

                    cellMonths.ForEach(x =>
                    {
                        x.CellValue = new CellValue(bktReport.Time.Month.ToString());
                        x.DataType = new EnumValue<CellValues>(CellValues.Number);
                    });

                    cellYears.ForEach(x =>
                    {
                        x.CellValue = new CellValue(bktReport.Time.Year.ToString());
                        x.DataType = new EnumValue<CellValues>(CellValues.Number);
                    });

                    #endregion

                    #region Report information fill

                    var rowIndex = 6;
                    while (true)
                    {
                        var cell = sheet.
                            Descendants<Cell>().FirstOrDefault(c => c.CellReference == $"A{rowIndex}");

                        if (cell == null)
                            break;

                        var iElementIndex = -1;
                        if (!int.TryParse(cell.CellValue.InnerText, out iElementIndex))
                            break;
                        
                        if (iElementIndex < 1 || iElementIndex > bktReport.Reports.Count)
                            break;

                        iElementIndex--;

                        for (var alphabet = 'C'; alphabet <= 'Z'; alphabet++)
                        {
                            var vCell = sheet.
                                Descendants<Cell>().FirstOrDefault(c => c.CellReference == $"{alphabet}{rowIndex}");

                            if (vCell == null)
                                continue;

                            var iIndex = alphabet - 'C';
                            var info = bktReport.Reports[iElementIndex][iIndex];
                            vCell.CellValue = new CellValue(info.Rain.ToString("F"));
                        }
                    }

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
    }
}