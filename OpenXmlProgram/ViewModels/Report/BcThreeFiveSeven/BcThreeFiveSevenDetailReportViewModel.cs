using System;

namespace OpenXmlProgram.ViewModels.Report.BcThreeFiveSeven
{
    public class BcThreeFiveSevenDetailReportViewModel
    {
        /// <summary>
        /// Total amount of rain.
        /// </summary>
        public double? Rr { get; set; }

        /// <summary>
        /// Air temperature.
        /// </summary>
        public double? T { get; set; }

        /// <summary>
        /// Maximum air temperature.
        /// </summary>
        public double? Tx { get; set; }

        /// <summary>
        /// Minimum air temperature.
        /// </summary>
        public double? Tn { get; set; }

        /// <summary>
        /// Air humidity.
        /// </summary>
        public double? U { get; set; }

        /// <summary>
        /// Minimum air humidity.
        /// </summary>
        public double? Un { get; set; }

        /// <summary>
        /// Air pressure of ocean.
        /// </summary>
        public double? Po { get; set; }

        /// <summary>
        /// Air pressure.
        /// </summary>
        public double? P { get; set; }

        /// <summary>
        /// Maximum air pressure.
        /// </summary>
        public double? Px { get; set; }

        /// <summary>
        /// Maximum ocean air pressure.
        /// </summary>
        public double? Pn { get; set; }

        public double? Bh { get; set; }

        public double? Dd { get; set; }

        public double? Ff { get; set; }

        public double? Dxdx2M { get; set; }

        public double? FxFx2M { get; set; }

        public double? Dxdx2S { get; set; }

        public double? Fxfx2S { get; set; }

        public double? Bxc { get; set; }

        public DateTime Time { get; set; }

    }
}