namespace ExcelDNAVSExtension
{
    class ExceptionHandler
    {
        public static void ShowException(System.Exception e)
        {
            System.Windows.MessageBox.Show(e.ToString(), "Excel-DNA Developer Tools exception.", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
        }
    }
}
