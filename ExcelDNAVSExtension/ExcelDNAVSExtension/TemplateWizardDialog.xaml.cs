using System.Windows;

namespace ExcelDNAVSExtension
{
    public partial class TemplateWizardDialog : Window
    {
        public TemplateWizardDialog(ProjectCreationOptions options)
        {
            InitializeComponent();

            this.options = options;
            includeRibbonCheckBox.IsChecked = options.includeRibbon;
            includeXMLSchemasCheckBox.IsChecked = options.includeXMLSchemas;
        }

        public ProjectCreationOptions GetOptions()
        {
            return options;
        }

        private void OnCreate(object sender, RoutedEventArgs args)
        {
            try
            {
                options.includeRibbon = includeRibbonCheckBox.IsChecked.Value;
                options.includeXMLSchemas = includeXMLSchemasCheckBox.IsChecked.Value;
                DialogResult = true;
            }
            catch (System.Exception e)
            {
                ExceptionHandler.ShowException(e);
            }
        }

        private ProjectCreationOptions options;
    }
}
