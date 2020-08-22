namespace ExcelDNAVSExtension
{
    [System.Runtime.InteropServices.Guid("7c62086f-14b6-45f0-a11b-31698fd43e99")]
    class VBAddinProjectFactory : AddinProjectFactory
    {
        public VBAddinProjectFactory(VSPackage package) : base(package, typeof(VBAddinProjectFactory).GUID)
        {
        }
    }
}
