using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;

namespace $rootnamespace$
{
    [ComVisible(true)]
public class $safeitemrootname$ : ExcelRibbon
{
    public override string GetCustomUI(string RibbonID)
{
    return $fileinputname$Resources.Ribbon;
}

public override object LoadImage(string imageId)
{
    // This will return the image resource with the name specified in the image='xxxx' tag
    return $fileinputname$Resources.ResourceManager.GetObject(imageId);
}

public void OnButtonPressed(IRibbonControl control)
{
    System.Windows.Forms.MessageBox.Show("Hello!");
}
}
}
