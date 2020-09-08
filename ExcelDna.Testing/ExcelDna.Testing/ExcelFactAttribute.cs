using System;
using Xunit.Sdk;
using ExcelDna.Testing;

namespace Xunit
{
    [AttributeUsage(AttributeTargets.Method, AllowMultiple = false)]
    [XunitTestCaseDiscoverer("ExcelDna.Testing." + nameof(ExcelFactDiscoverer), "ExcelDna.Testing")]
    public class ExcelFactAttribute : FactAttribute
    {
        public bool UseCOM { get; set; }
        public string Workbook { get; set; }
    }
}
