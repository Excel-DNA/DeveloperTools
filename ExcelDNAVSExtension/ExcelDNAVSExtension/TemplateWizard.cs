using System.Collections.Generic;
using Microsoft.VisualStudio.TemplateWizard;
using EnvDTE;

namespace ExcelDNAVSExtension
{
    public class TemplateWizard : IWizard
    {
        public void BeforeOpeningFile(ProjectItem projectItem)
        {
        }

        public void ProjectFinishedGenerating(Project project)
        {
            try
            {
                XmlSchemas.AddSchemasToProject(project);
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
            replacementsDictionary.Add("$exceldnaincluderibbon$", "true");
        }

        public bool ShouldAddProjectItem(string filePath)
        {
            return true;
        }
    }
}
