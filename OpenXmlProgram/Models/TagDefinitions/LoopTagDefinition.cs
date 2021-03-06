﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Mustache;

namespace OpenXmlProgram.Models.TagDefinitions
{
    public class LoopTagDefinition : ContentTagDefinition
    {
        #region Properties

        private const string StartIndexParameter = "start";
        private static readonly TagParameter StartIndex = new TagParameter(StartIndexParameter) { IsRequired = true };

        private const string EndIndexParameter = "end";
        private static readonly TagParameter EndIndex = new TagParameter(EndIndexParameter) { IsRequired = true };

        private const string OperatorParameter = "operator";
        private static readonly TagParameter Operator = new TagParameter(OperatorParameter) { IsRequired = true };

        private const string IncreasementParameter = "increasement";
        private static readonly TagParameter Increasement = new TagParameter(IncreasementParameter) {IsRequired = false};
        #endregion

        #region Constructor

        /// <summary>
        ///     Initializes a new instance of an EachTagDefinition.
        /// </summary>
        public LoopTagDefinition()
            : base("Loop")
        {
        }

        #endregion

        #region Methods

        /// <summary>
        ///     Gets whether the tag only exists within the scope of its parent.
        /// </summary>
        protected override bool GetIsContextSensitive()
        {
            return false;
        }

        /// <summary>
        ///     Gets the parameters that can be passed to the tag.
        /// </summary>
        /// <returns>The parameters.</returns>
        protected override IEnumerable<TagParameter> GetParameters()
        {
            return new[] { StartIndex, Operator, EndIndex, Increasement };
        }

        /// <summary>
        ///     Gets the context to use when building the inner text of the tag.
        /// </summary>
        /// <param name="writer">The text writer passed</param>
        /// <param name="keyScope">The current scope.</param>
        /// <param name="arguments">The arguments passed to the tag.</param>
        /// <param name="contextScope"></param>
        /// <returns>The scope to use when building the inner text of the tag.</returns>
        public override IEnumerable<NestedContext> GetChildContext(
            TextWriter writer,
            Scope keyScope,
            Dictionary<string, object> arguments,
            Scope contextScope)
        {
            var oStartIndex = arguments[StartIndexParameter];
            var oEndIndex = arguments[EndIndexParameter];
            var oOperator = arguments[OperatorParameter];
            

            var iStartIndex = Convert.ToInt32(oStartIndex);
            var iEndIndex = Convert.ToInt32(oEndIndex);
            var szOperator = Convert.ToString(oOperator);

            // Increasement detect.
            var increasement = 1;
            if (arguments.ContainsKey(IncreasementParameter))
            {
                var oIncreasement = arguments[IncreasementParameter];
                increasement = Convert.ToInt32(oIncreasement);
            }

            var index = iStartIndex;
            for (; ; index+= increasement)
            {
                if ("lower-equal".Equals(szOperator, StringComparison.InvariantCultureIgnoreCase))
                {
                    if (!(index <= iEndIndex))
                        break;
                }

                if ("lower-than".Equals(szOperator, StringComparison.InvariantCultureIgnoreCase))
                {
                    if (!(index < iEndIndex))
                        break;
                }

                if ("greater-than".Equals(szOperator, StringComparison.InvariantCultureIgnoreCase))
                {
                    if (!(index > iEndIndex))
                        break;
                }

                if ("greater-equal".Equals(szOperator, StringComparison.InvariantCultureIgnoreCase))
                {
                    if (!(index >= iEndIndex))
                        break;
                }
                
                var output = new
                {
                    pointer = index
                };
                var childContext = new NestedContext
                {
                    KeyScope = keyScope.CreateChildScope(output),
                    Writer = writer,
                    ContextScope = contextScope.CreateChildScope(output)
                };
                yield return childContext;
            }
        }

        /// <summary>
        ///     Gets the parameters that are used to create a new child context.
        /// </summary>
        /// <returns>The parameters that are used to create a new child context.</returns>
        public override IEnumerable<TagParameter> GetChildContextParameters()
        {
            return new[] { StartIndex, Operator, EndIndex };
        }

        #endregion
    }
}