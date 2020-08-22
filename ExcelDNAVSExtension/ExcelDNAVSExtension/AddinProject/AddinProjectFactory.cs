namespace ExcelDNAVSExtension
{
    class AddinProjectFactory : Microsoft.VisualStudio.Shell.Flavor.FlavoredProjectFactoryBase
    {
        public AddinProjectFactory(VSPackage package, System.Guid addItemTemplates)
        {
            this.package = package;
            this.addItemTemplates = addItemTemplates;
        }

        protected override object PreCreateForOuter(System.IntPtr outerProjectIUnknown)
        {
            return new AddinProject(package, addItemTemplates);
        }

        private Microsoft.VisualStudio.Shell.Package package;
        private System.Guid addItemTemplates;
    }
}
