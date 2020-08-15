using System.Collections.Generic;
using Microsoft.VisualStudio.TemplateWizard;
using EnvDTE;

namespace ExcelDNAVSExtension
{
    public class TemplateWizard : IWizard
    {
        public TemplateWizard()
        {
            options = new ProjectCreationOptions();
            options.includeRibbon = true;
            options.includeXMLSchemas = true;
        }

        public void BeforeOpeningFile(ProjectItem projectItem)
        {
        }

        public void ProjectFinishedGenerating(Project project)
        {
            try
            {
                if (options.includeXMLSchemas)
                    XmlSchemas.AddSchemasToProject(project);

                if (options.includeRibbon)
                    RunResxCustomTool(project);

                VSExceptionSettings.DisableLoaderLock();
            }
            catch (System.Exception e)
            {
                ExceptionHandler.ShowException(e);
            }
        }

        public void ProjectItemFinishedGenerating(ProjectItem projectItem)
        {
        }

        public void RunFinished()
        {
        }

        public void RunStarted(object automationObject, Dictionary<string, string> replacementsDictionary, WizardRunKind runKind, object[] customParams)
        {
            TemplateWizardDialog dialog = new TemplateWizardDialog(options);
            dialog.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen;
            if (dialog.ShowDialog().GetValueOrDefault())
            {
                options = dialog.GetOptions();
                replacementsDictionary.Add("$exceldnaincluderibbon$", options.includeRibbon ? "true" : "false");
            }
            else
            {
                throw new WizardBackoutException();
            }
        }

        public bool ShouldAddProjectItem(string filePath)
        {
            return true;
        }

        private void RunResxCustomTool(Project project)
        {
            foreach (ProjectItem projectItem in project.ProjectItems)
            {
                if (projectItem.Name.EndsWith(".resx"))
                {
                    VSLangProj.VSProjectItem vsProjectItem = projectItem.Object as VSLangProj.VSProjectItem;
                    vsProjectItem?.RunCustomTool();
                }
            }
        }

        private ProjectCreationOptions options;
    }
}
