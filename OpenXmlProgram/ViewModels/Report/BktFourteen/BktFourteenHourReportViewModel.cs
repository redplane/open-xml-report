using System;

namespace OpenXmlProgram.ViewModels.Report.BktFourteen
{
    public class BktFourteenHourReportViewModel
    {
        /// <summary>
        /// Total rain of hour.
        /// </summary>
        public double Rain { get; set; }

        /// <summary>
        /// Time during rain.
        /// </summary>
        public double RainObs { get; set; }

        /// <summary>
        /// Time of system.
        /// </summary>
        public DateTime Time { get; set; }
    }
}