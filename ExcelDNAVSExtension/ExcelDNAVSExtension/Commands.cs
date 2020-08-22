namespace ExcelDNAVSExtension
{
    class Commands
    {
        public Commands(Microsoft.VisualStudio.Shell.Package package)
        {
            System.IServiceProvider serviceProvider = package as System.IServiceProvider;
            mcs = (Microsoft.VisualStudio.Shell.OleMenuCommandService)serviceProvider.GetService(
                    typeof(System.ComponentModel.Design.IMenuCommandService));

            AddCommand(() => System.Diagnostics.Process.Start("https://docs.excel-dna.net"), cmdidDocs);
            AddCommand(() => System.Windows.MessageBox.Show("Excel-DNA Developer Tools v1.0.2.", "About", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information), cmdidAbout);
        }

        private void AddCommand(System.Action command, uint cmdid)
        {
            mcs.AddCommand(CreateMenuCommand((object sender, System.EventArgs args) =>
            {
                try
                {
                    command();
                }
                catch (System.Exception e)
                {
                    ExceptionHandler.ShowException(e);
                }
            }, cmdid));
        }

        private System.ComponentModel.Design.MenuCommand CreateMenuCommand(System.EventHandler onClick, uint id)
        {
            System.ComponentModel.Design.CommandID menuCommandID = new System.ComponentModel.Design.CommandID(CommandSet, (int)id);
            return new System.ComponentModel.Design.MenuCommand(onClick, menuCommandID);
        }

        private Microsoft.VisualStudio.Shell.OleMenuCommandService mcs;

        private static readonly System.Guid CommandSet = new System.Guid("a6a2b6bd-05ce-4df3-a1c2-a9fa1685c969");
        private const uint cmdidAbout = 0x100;
        private const uint cmdidDocs = 0x101;
    }
}
