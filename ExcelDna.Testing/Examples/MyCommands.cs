using Examples;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace test_ExcelAddin2
{
    public static class MyCommands
    {
        // We make a command macro that can be run by:
        // * Pressing the quick menu under the Add-ins tab
        // * Pressing the shortcut key Ctrl + Shift + D
        // * Typing the name into the Alt+F8 Macro dialog (add-in macros won't we shown on this list, though)
        [ExcelCommand(MenuName = "Example", MenuText = "Dump Data", ShortCut = "^D")]
        public static void DumpData()
        {
            System.Windows.Forms.MessageBox.Show(Utils.GetVersion());
        }
    }
}
