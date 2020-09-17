using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using Xunit;

namespace Examples
{
    public class InProcessTests
    {
        [ExcelFact]
        public void GetExcelVersion()
        {
            Assert.Equal("16.0", Utils.GetVersion());
        }

        [ExcelFact(Workbook = "")]
        public void SetCellValue()
        {
            Range targetRange = ExcelDna.Testing.Util.Workbook.Sheets[1].Range["A1:C2"];

            object[,] newValues = new object[,] { { "One", 2, "Three" }, { true, DateTime.Now, "" } };
            targetRange.Value = newValues;

            var cell = new ExcelReference(0, 0);
            Assert.Equal("One", cell.GetValue().ToString());
        }

        [ExcelFact]
        public void ApplicationVersion()
        {
            Assert.Equal("16.0", ExcelDna.Testing.Util.Application.Version);
        }

        [ExcelFact(Workbook = "MrExcel.xlsx")]
        public void PreCreatedWorkbook()
        {
            Range cell = ExcelDna.Testing.Util.Workbook.Sheets[1].Range["A2:A2"];
            Assert.Equal("Red Ford Truck", cell.Value.ToString());
        }

        [ExcelFact(Workbook = "", XLL = @"..\..\..\ExampleAddin\bin\Debug\ExampleAddin-AddIn")]
        public void ClickRibbonButton()
        {
            ExcelDna.Testing.Automation.ClickRibbonButton("ExampleAddin", "My Button");

            var cell = new ExcelReference(0, 0);
            Assert.Equal("One", cell.GetValue().ToString());
        }
    }
}
