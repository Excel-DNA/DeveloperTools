using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;

namespace $safeprojectname$
{
    [ComVisible(true)]
public class MyRibbon : ExcelRibbon
{
    public override string GetCustomUI(string RibbonID)
    {
        using (System.IO.Stream stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(
            "$safeprojectname$." + "MyRibbon.xml"))
        {
            using (System.IO.StreamReader reader = new System.IO.StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }
    }

    public void OnButtonPressed(IRibbonControl control)
    {
        System.Windows.Forms.MessageBox.Show("Hello!");
    }
}
}
