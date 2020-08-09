using System.IO;

namespace ExcelDNAVSExtension
{
    class XmlSchemas
    {
        public static void AddSchemasToProject(EnvDTE.Project project)
        {
            string schemasDir = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "XmlSchemas");
            foreach (string file in Directory.GetFiles(schemasDir))
            {
                if (!IsSchemaInCache(Path.GetFileName(file), project.DTE.LocaleID))
                    project.ProjectItems.AddFromFileCopy(file);
            }
        }

        private static bool IsSchemaInCache(string fileName, int locale)
        {
            string schemaCacheDir = Path.Combine(Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName), @"..\..\Xml\Schemas");
            return File.Exists(Path.Combine(schemaCacheDir, fileName)) || File.Exists(Path.Combine(schemaCacheDir, locale.ToString(), fileName));
        }
    }
}
