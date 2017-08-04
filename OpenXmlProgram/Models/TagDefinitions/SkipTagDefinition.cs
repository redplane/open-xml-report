using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Mustache;

namespace OpenXmlProgram.Models.TagDefinitions
{
    public class SkipTagDefinition : ContentTagDefinition
    {
        #region Constructor

        /// <summary>
        ///     Initiate tag definition with parameters.
        /// </summary>
        public SkipTagDefinition()
            : base("skip")
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
            return new[] {
                new TagParameter(Collection) { IsRequired = true },
                new TagParameter(Skip) { IsRequired = true },
                new TagParameter(Take) { IsRequired = true }
            };
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
            // Parameter parse.
            var collection = arguments[Collection];
            var skip = arguments[Skip];
            var take = arguments[Take];
            
            // Cast the collection.
            var enumerable = collection as IList;
            if (enumerable == null)
                return null;

            // Numeric conversion.
            var iSkip = Convert.ToInt32(skip);
            var iTake = Convert.ToInt32(take);

            // Skip and take.
            var items = from x in enumerable
                        

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
        private const string Collection = "Collection";

        /// <summary>
        ///     Skip value.
        /// </summary>
        private const string Skip = "Skip";

        /// <summary>
        /// Take value.
        /// </summary>
        private const string Take = "Take";


        #endregion
    }
}