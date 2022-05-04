using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit.Abstractions;

namespace ExcelDna.Testing
{
    internal class RunnerUtil
    {
        public static string TestAssemblyDirectory(ITestAssembly testAssembly, IEnumerable<ExcelTestCase> testCases)
        {
            string result = Path.GetDirectoryName(testAssembly.Assembly.AssemblyPath);
            if (string.IsNullOrEmpty(result))
            {
                ExcelTestCase test = testCases?.FirstOrDefault();
                result = Path.GetDirectoryName((test?.Method as Xunit.Sdk.ReflectionMethodInfo)?.MethodInfo?.DeclaringType?.Assembly?.Location);
                if (result == null)
                    result = "";
            }

            return result;
        }
    }
}
