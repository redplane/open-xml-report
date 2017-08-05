using System.Collections.Generic;
using OpenXmlProgram.Models;

namespace OpenXmlProgram.ViewModels.Report.BcThreeFiveSeven
{
    public class BcThreeFiveSevenReportViewModel
    {
        /// <summary>
        /// Station information.
        /// </summary>
        public Station Station { get; set; }

        /// <summary>
        /// Report details.
        /// </summary>
        public IEnumerable<BcThreeFiveSevenDetailReportViewModel> Details { get; set; }
    }
}