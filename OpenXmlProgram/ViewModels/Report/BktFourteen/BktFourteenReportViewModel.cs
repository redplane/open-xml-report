using System;
using System.Collections.Generic;
using OpenXmlProgram.Models;

namespace OpenXmlProgram.ViewModels.Report.BktFourteen
{
    public class BktFourteenReportViewModel
    {
        /// <summary>
        /// Station information.
        /// </summary>
        public Station Station { get; set; }

        /// <summary>
        /// List of reports.
        /// </summary>
        public IList<IList<BktFourteenHourReportViewModel>> Reports { get; set; }

        /// <summary>
        /// Time of report.
        /// </summary>
        public DateTime Time { get; set; }
    }
}