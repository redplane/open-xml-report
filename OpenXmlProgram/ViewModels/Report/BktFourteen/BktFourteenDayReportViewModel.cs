using System;
using System.Collections.Generic;

namespace OpenXmlProgram.ViewModels.Report.BktFourteen
{
    public class BktFourteenDayReportViewModel
    {
        /// <summary>
        /// Reports of minute
        /// </summary>
        public IList<BktFourteenHourReportViewModel> Hours { get; set; }

        
        /// <summary>
        /// Time of hour.
        /// </summary>
        public DateTime Time { get; set; }
    }
}