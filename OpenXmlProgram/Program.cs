
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
            compiler.RegisterTag(new NumericCompareTagDefinition(), true);
            compiler.RegisterTag(new LookupTagDefinition(), true);
            compiler.RegisterTag(new LoopTagDefinition(), true);
            
            //ExportBktTen(path, compiler);
            ExportBktFourteen(path, compiler);
        }

        #region Methods

        /// <summary>
        /// Export bkt-10 report.
        /// </summary>
        /// <param name="path"></param>
        /// <param name="formatCompiler"></param>
        private static void ExportBktTen(string path, FormatCompiler formatCompiler)
        {
            // Path to which file should be exported.
            var output = Path.Combine(path, $"{DateTime.Now.ToString("yy-MM-dd HH-mm-ss")}.xls");
            var input = Path.Combine(path, "Files/bkt-10.xml");

            // Template file.
            var szWindPath = Path.Combine(path, "Data/bkt-10.json");
            var szText = File.ReadAllText(szWindPath);
            var szInput = File.ReadAllText(input);

            // Deserialize text.
            var result = JsonConvert.DeserializeObject<BktTenReportViewModel>(szText);
            foreach (var detail in result.Details)
            {
                if (detail.Reports.Count > 22)
                    continue;

                // Initiate new reports list.
                var reports = new Wind[24];

                for (var iReportIndex = 0; iReportIndex < detail.Reports.Count; iReportIndex++)
                {
                    var report = detail.Reports[iReportIndex];
                    reports[report.DateTime.Hour] = report;
                }

                detail.Reports = reports;
            }

            var generator = formatCompiler.Compile(szInput);
            var szXls = generator.Render(result);

            File.WriteAllText(output, szXls, Encoding.UTF8);
            Process.Start(output);
        }

        /// <summary>
        /// Export bkt-14 report.
        /// </summary>
        /// <param name="path"></param>
        /// <param name="formatCompiler"></param>
        private static void ExportBktFourteen(string path, FormatCompiler formatCompiler)
        {
            // Path to which file should be exported.
            var output = Path.Combine(path, $"{DateTime.Now.ToString("yy-MM-dd HH-mm-ss")}.xls");
            var input = Path.Combine(path, "Files/bkt-14.xml");

            // Template file.
            var szRainPath = Path.Combine(path, "Data/bkt-14.json");
            var szText = File.ReadAllText(szRainPath);
            var szInput = File.ReadAllText(input);

            // Deserialize text.
            var result = JsonConvert.DeserializeObject<BktFourteenReportViewModel>(szText);
            //foreach (var detail in result.Details)
            //{
            //    if (detail.Reports.Count > 22)
            //        continue;

            //    // Initiate new reports list.
            //    var reports = new Wind[24];

            //    for (var iReportIndex = 0; iReportIndex < detail.Reports.Count; iReportIndex++)
            //    {
            //        var report = detail.Reports[iReportIndex];
            //        reports[report.DateTime.Hour] = report;
            //    }

            //    detail.Reports = reports;
            //}

            var generator = formatCompiler.Compile(szInput);
            var szXls = generator.Render(result);

            File.WriteAllText(output, szXls, Encoding.UTF8);
            Process.Start(output);
        }

        #endregion
    }
}