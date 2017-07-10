using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Mustache;

namespace OpenXmlProgram.Models.TagDefinitions
{
    public class FindItemInArrayTagDefinition : ContentTagDefinition
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
        public FindItemInArrayTagDefinition()
            : base("FindItemInArray")
        { }

        #endregion

        public override IEnumerable<TagParameter> GetChildContextParameters()
        {
            return new TagParameter[0];
        }

        /// <summary>
        /// Get list of parameters.
        /// </summary>
        /// <returns></returns>
        protected override IEnumerable<TagParameter> GetParameters()
        {
            return new[] { new TagParameter(Source) { IsRequired = true }, new TagParameter(Target) { IsRequired = true } };
        }

        /// <summary>
        /// Unknown.
        /// </summary>
        /// <returns></returns>
        protected override bool GetIsContextSensitive()
        {
            return false;
        }
        
        public override IEnumerable<NestedContext> GetChildContext(TextWriter writer, Scope keyScope, Dictionary<string, object> arguments, Scope contextScope)
        {
            var rawCollection = arguments[Source];
            var rawIndex = arguments[Target];

            var enumerable = rawCollection as IList;
            if (enumerable == null)
                return null;
            
            var iIndex = Convert.ToInt32(rawIndex);
            var item = enumerable[iIndex];
            var childContext = new NestedContext()
            {
                KeyScope = keyScope.CreateChildScope(item),
                Writer = writer,
                ContextScope = contextScope.CreateChildScope(item),
            };

            return new List<NestedContext>() { childContext };
        }
    }
}