using ExcelDna.Integration;
using System.IO.Pipes;

namespace ExcelAgent
{
    public class Addin : IExcelAddIn
    {
        public void AutoOpen()
        {
            new System.Threading.Thread(StartRPC).Start();

            try
            {
                ExcelDna.Testing.ExcelStartupEvent.Set();
            }
            catch (System.Threading.WaitHandleCannotBeOpenedException)
            {
            }
        }

        public void AutoClose()
        {
        }

#pragma warning disable VSTHRD100
#if NET10_0_OR_GREATER
        [System.Diagnostics.CodeAnalysis.DynamicDependency(System.Diagnostics.CodeAnalysis.DynamicallyAccessedMemberTypes.PublicMethods, typeof(ExcelDna.Testing.Remote.RemoteExcel))]
        [System.Diagnostics.CodeAnalysis.DynamicDependency(System.Diagnostics.CodeAnalysis.DynamicallyAccessedMemberTypes.All, typeof(ExcelDna.Testing.Remote.RemoteExcel))]
#endif
        private async void StartRPC()
        {
            try
            {
                var stream = new NamedPipeServerStream("ExcelDna.Testing", PipeDirection.InOut, NamedPipeServerStream.MaxAllowedServerInstances, PipeTransmissionMode.Byte, PipeOptions.Asynchronous);
                await stream.WaitForConnectionAsync();
                var jsonRpc = StreamJsonRpc.JsonRpc.Attach(stream, new ExcelDna.Testing.Remote.RemoteExcel());
                await jsonRpc.Completion;
            }
            catch
            {
            }
        }
#pragma warning restore VSTHRD100
    }
}
