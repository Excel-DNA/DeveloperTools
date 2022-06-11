using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;

namespace ExampleAddin
{
    [ComVisible(true)]
    public class MyRibbon : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return RibbonResources.Ribbon;
        }

        public override object LoadImage(string imageId)
        {
            // This will return the image resource with the name specified in the image='xxxx' tag
            return RibbonResources.ResourceManager.GetObject(imageId);
        }

        public void OnButtonPressed(IRibbonControl control)
        {
            Application app = (Application)ExcelDnaUtil.Application;
            Range targetRange = (app.ActiveSheet as Worksheet).Range["A1:C2"];

            object[,] newValues = new object[,] { { "One", 2, "Three" }, { true, System.DateTime.Now, "" } };
            targetRange.Value = newValues;
        }
    }
}
