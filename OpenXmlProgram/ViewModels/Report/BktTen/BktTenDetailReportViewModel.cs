using System;
using System.Collections.Generic;
using OpenXmlProgram.Models;

namespace OpenXmlProgram.ViewModels.Report.BktTen
{
    public class BktTenDetailReportViewModel
    {
        #region Properties

        /// <summary>
        /// List of reports.
        /// </summary>
        public IList<Wind> Reports { get; set; }

        /// <summary>
        /// Time of report.
        /// </summary>
        public DateTime Time { get; set; }

        #endregion
    }
}