using System;
using System.Collections.Generic;
using OpenXmlProgram.Models;

namespace OpenXmlProgram.ViewModels.Report.BktFourteen
{
    public class BktFourteenReportViewModel
    {
        #region Properties

        /// <summary>
        ///     Station information.
        /// </summary>
        public Station Station { get; set; }

        /// <summary>
        ///     List of reports.
        /// </summary>
        public IList<BktFourteenDayReportViewModel> Reports { get; set; }

        /// <summary>
        /// Total rain of hour.
        /// </summary>
        public double[] TotalRains { get; set; }

        /// <summary>
        /// Total rain obs.
        /// </summary>
        public double[] TotalObservations { get; set; }

        /// <summary>
        ///     Time which report is about.
        /// </summary>
        public DateTime Time { get; set; }

        #endregion
    }
}