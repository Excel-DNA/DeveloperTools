using Microsoft.VisualStudio.Shell.Flavor;

namespace ExcelDNAVSExtension
{
    class AddinProject : FlavoredProjectBase
    {
        public AddinProject(Microsoft.VisualStudio.Shell.Package package, System.Guid addItemTemplates)
        {
            this.package = package;
            this.addItemTemplates = addItemTemplates;
        }

        protected override void SetInnerProject(System.IntPtr innerIUnknown)
        {
            if (this.serviceProvider == null)
                this.serviceProvider = package;
            base.SetInnerProject(innerIUnknown);
        }

        protected override System.Guid GetGuidProperty(uint itemId, int propId)
        {
            if (propId == (int)Microsoft.VisualStudio.Shell.Interop.__VSHPROPID2.VSHPROPID_AddItemTemplatesGuid)
            {
                return addItemTemplates;
            }
            return base.GetGuidProperty(itemId, propId);
        }

        private Microsoft.VisualStudio.Shell.Package package;
        private System.Guid addItemTemplates;
    }
}
