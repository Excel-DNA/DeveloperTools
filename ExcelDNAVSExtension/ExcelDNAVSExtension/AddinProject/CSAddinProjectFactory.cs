namespace ExcelDNAVSExtension
{
    [System.Runtime.InteropServices.Guid("211e647a-3b33-4448-bc03-fac6d799aa09")]
    class CSAddinProjectFactory : AddinProjectFactory
    {
        public CSAddinProjectFactory(VSPackage package) : base(package, typeof(CSAddinProjectFactory).GUID)
        {
        }
    }
}
