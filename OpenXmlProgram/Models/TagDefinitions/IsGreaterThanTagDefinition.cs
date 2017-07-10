using System;
using System.Collections.Generic;
using Mustache;

namespace OpenXmlProgram.Models.TagDefinitions
{
    public class IsGreaterThanTagDefinition: ContentTagDefinition
    {
        #region Properties

        /// <summary>
        /// Value of source.
        /// </summary>
        private const string Source = "Source";

        /// <summary>
        /// Target value.
        /// </summary>
        private const string Target = "Target";

        #endregion

        #region Constructor

        /// <summary>
        /// Initiate tag definition with parameters.
        /// </summary>
        public IsGreaterThanTagDefinition()
            : base("IsGreaterThan")
        { }

        #endregion

        public override IEnumerable<TagParameter> GetChildContextParameters()
        {
            return new TagParameter[0];
        }

        public override bool ShouldGeneratePrimaryGroup(Dictionary<string, object> arguments)
        {
            var source = arguments[Source];
            var target = arguments[Target];

            return IsConditionSatisfied(source, target);
        }

        /// <summary>
        /// Get list of parameters.
        /// </summary>
        /// <returns></returns>
        protected override IEnumerable<TagParameter> GetParameters()
        {
            return new [] { new TagParameter(Source) { IsRequired = true }, new TagParameter(Target) { IsRequired = true } };
        }

        /// <summary>
        /// Unknown.
        /// </summary>
        /// <returns></returns>
        protected override bool GetIsContextSensitive()
        {
            return false;
        }

        /// <summary>
        /// Check whether condition is satisfied.
        /// </summary>
        /// <param name="source"></param>
        /// <param name="target"></param>
        /// <returns></returns>
        private bool IsConditionSatisfied(object source, object target)
        {
            // Condition is invalid.
            if (source == null || target == null)
                return true;

            var dSource = Convert.ToDouble(source);
            var dTarget = Convert.ToDouble(target);

            return dSource > dTarget;
        }
    }
}