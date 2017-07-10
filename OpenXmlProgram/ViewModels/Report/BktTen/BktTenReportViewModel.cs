using System.Collections.Generic;
using OpenXmlProgram.Models;

namespace OpenXmlProgram.ViewModels.Report.BktTen
{
    public class BktTenReportViewModel
    {
        #region Properties

        /// <summary>
        /// Station information.
        /// </summary>
        public Station Station { get; set; }

        /// <summary>
        /// Detailed report.
        /// </summary>
        public IEnumerable<BktTenDetailReportViewModel> Details { get; set; }

        #endregion
    }
}