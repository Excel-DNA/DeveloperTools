using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using Xunit;

[assembly: Xunit.TestFramework("Xunit.ExcelTestFramework", "ExcelDna.Testing")]

namespace Examples
{
    public class UtilsTest
    {
        [Fact]
        public void RegularTest()
        {
            Assert.Equal(4, 2 + 2);
        }

        [ExcelFact]
        public void GetExcelVersion()
        {
            Assert.Equal("16.0", Utils.GetVersion());
        }

        [ExcelFact]
        public void SetCellValue()
        {
            Application app = (Application)ExcelDnaUtil.Application;
            var newBook = app.Workbooks.Add();

            Range targetRange = newBook.Sheets[1].Range["A1:C2"];

            object[,] newValues = new object[,] { { "One", 2, "Three" }, { true, DateTime.Now, "" } };
            targetRange.Value = newValues;

            var cell = new ExcelReference(0, 0);
            Assert.Equal("One", cell.GetValue().ToString());

            newBook.Close(false);
        }

        [ExcelFact(UseCOM = true)]
        public void COMGetExcelVersion()
        {
            Assert.Equal("16.0", ExcelDna.Testing.COMUtil.Application.Version);
        }

        [ExcelFact(UseCOM = true)]
        public void COMFunctionSayHello()
        {
            var newBook = ExcelDna.Testing.COMUtil.Application.Workbooks.Add();
            try
            {
                Range functionRange = newBook.Sheets[1].Range["B1:B1"];
                functionRange.Formula = "=SayHello(\"world\")";
                Assert.Equal("Hello world", functionRange.Value.ToString());
            }
            finally
            {
                newBook.Close(false);
            }
        }
    }
}
