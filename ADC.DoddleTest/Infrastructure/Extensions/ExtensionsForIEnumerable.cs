using System.Collections.Generic;
using ADC.DoddleTest.Infrastructure.Builder;

namespace ADC.DoddleTest.Infrastructure.Extensions
{
    public static class ExtensionsForIEnumerable
    {
        public static ExportBuilder<TModel> Export<TModel>(this IEnumerable<TModel> models) where TModel : class
        {
            return ExportBuilder<TModel>.Create(models);
        }
    }
}
