
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Xml.Xsl;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Mustache;
using Newtonsoft.Json;
using OpenXmlProgram.Models;
using OpenXmlProgram.Models.TagDefinitions;
using OpenXmlProgram.ViewModels.Report.BktFourteen;
using OpenXmlProgram.ViewModels.Report.BktTen;

namespace OpenXmlProgram
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            // Find application directory.
            var path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            // Initiate mustach-sharp compiler.
            var compiler = new FormatCompiler();
            compiler.RegisterTag(new IsGreaterThanTagDefinition(), true);
            compiler.RegisterTag(new IsSmallerThanTagDefinition(), true);
            compiler.RegisterTag(new LookupTagDefinition(), true);
            compiler.RegisterTag(new LoopTagDefinition(), true);

#if !USING_TXT
            // Path to which file should be exported.
            var output = Path.Combine(path, $"{DateTime.Now.ToString("yy-MM-dd HH-mm-ss")}.xls");
            var input = Path.Combine(path, "Files/bkt-10.xml");

            // Template file.
            var szWindPath = Path.Combine(path, "Data/bkt-10.json");
            var szText = File.ReadAllText(szWindPath);
            var szInput = File.ReadAllText(input);

            // Deserialize text.
            var result = JsonConvert.DeserializeObject<BktTenReportViewModel>(szText);

            

            var generator = compiler.Compile(szInput);
            var szXls = generator.Render(result);

            File.WriteAllText(output, szXls, Encoding.UTF8);
            Process.Start(output);
#endif

#if USING_TXT
            // Path to which file should be exported.
            var output = Path.Combine(path, $"bkt-10-{DateTime.Now.ToString("yy-MM-dd HH-mm-ss")}.txt");
            var input = Path.Combine(path, "Files/bkt-10.txt");

            // Template file.
            var szWindPath = Path.Combine(path, "Data/bkt-10.json");
            var szText = File.ReadAllText(szWindPath);
            var szInput = File.ReadAllText(input);

            // Deserialize text.
            var result = JsonConvert.DeserializeObject<BktTenReportViewModel>(szText);
            
            var generator = compiler.Compile(szInput);
            var szXls = generator.Render(result);

            File.WriteAllText(output, szXls, Encoding.UTF8);
            Process.Start(output);
#endif
        }

        protected static void AddPartXml(OpenXmlPart part, string xml)
        {
            using (var stream = part.GetStream())
            {
                var buffer = (new UTF8Encoding()).GetBytes(xml);
                stream.Write(buffer, 0, buffer.Length);
            }
        }

        /// <summary>
        ///     Find list of sheets by using specific conditions.
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
        ///     Find worksheet part by using sheet information.
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
        ///     Find rows in worksheet by using specific conditions.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        private static IEnumerable<Row> FindRows(Worksheet worksheet)
        {
            return worksheet.Elements<SheetData>().SelectMany(x => x.Elements<Row>());
        }

        /// <summary>
        ///     Find rows in worksheet by using specific conditions.
        /// </summary>
        /// <param name="worksheets"></param>
        /// <param name="condition"></param>
        /// <returns></returns>
        private static IEnumerable<Row> FindRows(IEnumerable<Worksheet> worksheets, Func<Row, bool> condition)
        {
            return worksheets.SelectMany(worksheet => worksheet.Elements<SheetData>()
                .SelectMany(x => x.Elements<Row>().Where(condition)));
        }

        /// <summary>
        ///     Find cells from collection of rows.
        /// </summary>
        /// <param name="rows"></param>
        /// <returns></returns>
        private static IEnumerable<Cell> FindCells(IEnumerable<Row> rows)
        {
            return rows.SelectMany(x => x.Elements<Cell>());
        }


        /// <summary>
        ///     Find cells by using specific conditions from collection of rows.
        /// </summary>
        /// <param name="rows"></param>
        /// <param name="condition"></param>
        /// <returns></returns>
        private static IEnumerable<Cell> FindCells(IEnumerable<Row> rows, Func<Cell, bool> condition)
        {
            return rows.SelectMany(x => x.Elements<Cell>().Where(condition));
        }

        /// <summary>
        ///     Find report information.
        /// </summary>
        /// <returns></returns>
        private static BktFourteenReportViewModel FindReport()
        {
            var path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            path = Path.Combine(path, "Data/bkt-14.json");

            var text = File.ReadAllText(path);
            //return JsonConvert.DeserializeObject<BktFourteenReportViewModel>(text);
            return null;
        }
        

        /// <summary>
        ///     Find cell at a specific position.
        /// </summary>
        /// <param name="rows"></param>
        /// <param name="column"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private static Cell FindCell(IEnumerable<Row> rows, string column, uint row)
        {
            return FindCells(rows).FirstOrDefault(c => string.Compare
                                                       (c.CellReference.Value, $"{column}{row}",
                                                           StringComparison.InvariantCultureIgnoreCase) == 0);
        }

        private static Cell InsertCell(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            var worksheet = worksheetPart.Worksheet;
            var sheetData = worksheet.GetFirstChild<SheetData>();
            var cellReference = columnName + rowIndex;
            Row row;
            if (sheetData.Elements<Row>().Count(r => r.RowIndex == rowIndex) != 0)
            {
                row = sheetData.Elements<Row>().First(r => r.RowIndex == rowIndex);
            }
            else
            {
                row = new Row {RowIndex = rowIndex};
                sheetData.Append(row);
            }
            if (row.Elements<Cell>().Any(c => c.CellReference.Value == columnName + rowIndex))
            {
                return row.Elements<Cell>().First(c => c.CellReference.Value == cellReference);
            }
            Cell refCell = null;
            foreach (var cell in row.Elements<Cell>())
                if (string.Compare(cell.CellReference.Value, cellReference, StringComparison.OrdinalIgnoreCase) > 0)
                {
                    refCell = cell;
                    break;
                }
            var newCell = new Cell {CellReference = cellReference};
            row.InsertBefore(newCell, refCell);
            worksheet.Save();
            return newCell;
        }
    }
}