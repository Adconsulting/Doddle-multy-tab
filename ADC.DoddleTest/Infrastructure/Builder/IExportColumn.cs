using System;

namespace ADC.DoddleTest.Infrastructure.Builder
{
    public interface IExportColumn<TModel> where TModel : class
    {
        string Header { get; }
        Func<TModel, Object> Display { get; }
    }
}
