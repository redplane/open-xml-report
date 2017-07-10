using System;
using System.Collections.Generic;
using Mustache;

namespace OpenXmlProgram.Models.TagDefinitions
{
    public class NumericCompareTagDefinition : ContentTagDefinition
    {
        #region Properties

        /// <summary>
        /// Value of source.
        /// </summary>
        private const string Source = "Source";

        /// <summary>
        /// Operator which is used for comparision.
        /// </summary>
        private const string Operator = "Operator";

        /// <summary>
        /// Target value.
        /// </summary>
        private const string Target = "Target";

        #endregion

        #region Constructor

        /// <summary>
        /// Initiate tag definition with parameters.
        /// </summary>
        public NumericCompareTagDefinition()
            : base("NumericCompare")
        { }

        #endregion

        #region Methods

        /// <summary>
        /// Get child context parameter.
        /// </summary>
        /// <returns></returns>
        public override IEnumerable<TagParameter> GetChildContextParameters()
        {
            return new TagParameter[0];
        }

        /// <summary>
        /// Should generate primary group
        /// </summary>
        /// <param name="arguments"></param>
        /// <returns></returns>
        public override bool ShouldGeneratePrimaryGroup(Dictionary<string, object> arguments)
        {
            var source = arguments[Source];
            var target = arguments[Target];
            var oOperator = arguments[Operator];

            // Condition is invalid.
            if (source == null || target == null)
                return true;

            var dSource = Convert.ToDouble(source);
            var dTarget = Convert.ToDouble(target);
            var szOperator = Convert.ToString(oOperator);

            if (dSource == 11)
            {
                var a = 1;
            }
                

            if ("lower-than".Equals(szOperator, StringComparison.InvariantCultureIgnoreCase))
                return dSource < dTarget;
            if ("lower-equal".Equals(szOperator, StringComparison.InvariantCultureIgnoreCase))
                return dSource <= dTarget;
            if ("equal".Equals(szOperator, StringComparison.InvariantCultureIgnoreCase))
                return dSource.Equals(dTarget);
            if ("greater-equal".Equals(szOperator, StringComparison.InvariantCultureIgnoreCase))
                return dSource >= dTarget;
            if ("greater".Equals(szOperator, StringComparison.InvariantCultureIgnoreCase))
                return dSource > dTarget;
            if ("different".Equals(szOperator, StringComparison.InvariantCultureIgnoreCase))
                return !dSource.Equals(dTarget);

            throw new Exception("Operator must be specified.");
        }

        /// <summary>
        /// Get list of parameters.
        /// </summary>
        /// <returns></returns>
        protected override IEnumerable<TagParameter> GetParameters()
        {
            return new[] { new TagParameter(Source) { IsRequired = true }, new TagParameter(Operator) { IsRequired = true }, new TagParameter(Target) { IsRequired = true } };
        }

        /// <summary>
        /// Unknown.
        /// </summary>
        /// <returns></returns>
        protected override bool GetIsContextSensitive()
        {
            return false;
        }

        #endregion
    }
}