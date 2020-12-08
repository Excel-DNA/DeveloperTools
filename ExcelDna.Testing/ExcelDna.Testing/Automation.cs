using System.Windows.Automation;

namespace ExcelDna.Testing
{
    public static class Automation
    {
        public static void ClickRibbonButton(string tabLabel, string buttonLabel)
        {
            AutomationElement ribbonTabs = AutomationElement.FromHandle(ExcelDna.Integration.ExcelDnaUtil.WindowHandle).FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Ribbon Tabs"));
            if (ribbonTabs == null)
                throw new System.Exception("Can't find Ribbon Tabs.");

            AutomationElement tabItem = ribbonTabs.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, tabLabel));
            if (tabItem == null)
                throw new System.Exception($"Can't find tab {tabLabel}.");
            (tabItem.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern).Select();

            AutomationElement ribbon = TreeWalker.ControlViewWalker.GetParent(ribbonTabs);
            AutomationElement lowerRibbon = ribbon.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "Lower Ribbon"));
            if (lowerRibbon == null)
                throw new System.Exception("Can't find Lower Ribbon.");

            AutomationElement button = lowerRibbon.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, buttonLabel));
            if (button == null)
                throw new System.Exception($"Can't find button {buttonLabel}.");
            (button.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern).Invoke();
            System.Windows.Forms.Application.DoEvents(); // Forcing the Invoke message to be processed now.
        }

        public static void Wait(int millisecondsTimeout)
        {
            WaitFor(() => false, millisecondsTimeout);
        }

        public static void WaitFor(System.Func<bool> condition, int millisecondsTimeout)
        {
            System.DateTime start = System.DateTime.UtcNow;
            while ((System.DateTime.UtcNow - start).TotalMilliseconds < millisecondsTimeout && !condition())
            {
                System.Windows.Forms.Application.DoEvents();
                System.Threading.Thread.Sleep(10);
            }
        }
    }
}
