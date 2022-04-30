using System.Collections.Generic;
using Xunit.Abstractions;

namespace ExcelDna.Testing.Remote
{
    internal class TestFrameworkOptions : ITestFrameworkExecutionOptions
    {
        readonly Dictionary<string, object> properties = new Dictionary<string, object>();

        public TValue GetValue<TValue>(string name, TValue defaultValue)
        {
            object result;
            if (properties.TryGetValue(name, out result))
                return (TValue)result;

            return defaultValue;
        }

        public TValue GetValue<TValue>(string name)
        {
            return GetValue(name, default(TValue));
        }

        public void SetValue<TValue>(string name, TValue value)
        {
            properties[name] = value;
        }
    }
}
