using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Mustache;

namespace OpenXmlProgram.Models.TagDefinitions
{
    public class DateFormatTagDefinition : ContentTagDefinition
    {
        #region Constructor

        /// <summary>
        ///     Initiate tag definition with parameters.
        /// </summary>
        public DateFormatTagDefinition()
            : base("FormatDate")
        {
        }

        #endregion

        /// <summary>
        ///     Get list of parameters in child context.
        /// </summary>
        /// <returns></returns>
        public override IEnumerable<TagParameter> GetChildContextParameters()
        {
            return new TagParameter[0];
        }

        /// <summary>
        ///     Get list of parameters.
        /// </summary>
        /// <returns></returns>
        protected override IEnumerable<TagParameter> GetParameters()
        {
            return new[] { new TagParameter(Date) { IsRequired = true }, new TagParameter(Format) { IsRequired = true } };
        }

        /// <summary>
        ///     Unknown.
        /// </summary>
        /// <returns></returns>
        protected override bool GetIsContextSensitive()
        {
            return false;
        }

        /// <summary>
        ///     Get child context.
        /// </summary>
        /// <param name="writer"></param>
        /// <param name="keyScope"></param>
        /// <param name="arguments"></param>
        /// <param name="contextScope"></param>
        /// <returns></returns>
        public override IEnumerable<NestedContext> GetChildContext(TextWriter writer, Scope keyScope,
            Dictionary<string, object> arguments, Scope contextScope)
        {
            var rawCollection = arguments[Date];
            var rawIndex = arguments[Format];

            var enumerable = rawCollection as IList;
            if (enumerable == null)
                return null;

            var iIndex = Convert.ToInt32(rawIndex);
            object item;
            try
            {
                item = enumerable[iIndex];
            }
            catch
            {
                item = null;
            }

            var childContext = new NestedContext
            {
                KeyScope = keyScope.CreateChildScope(item),
                Writer = writer,
                ContextScope = contextScope.CreateChildScope(item)
            };

            return new List<NestedContext> { childContext };
        }

        #region Properties

        /// <summary>
        ///     Value of source.
        /// </summary>
        private const string Date = "Date";

        /// <summary>
        ///     Format value.
        /// </summary>
        private const string Format = "Format";

        #endregion
    }
}