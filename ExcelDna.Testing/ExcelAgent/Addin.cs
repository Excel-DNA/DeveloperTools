using ExcelDna.Integration;
using System.Collections;
using System.Runtime.Remoting;
using System.Runtime.Remoting.Channels;
using System.Runtime.Remoting.Channels.Ipc;
using System.Runtime.Serialization.Formatters;
//using ExcelDna.Registration;
using System.Windows;

namespace ExcelAgent
{
    public class Addin : IExcelAddIn
    {
        public void AutoOpen()
        {
            channel = RegisterIpcChannel("ExcelAgent", "xxx1000", false);
            RemotingConfiguration.RegisterWellKnownServiceType(typeof(ExcelDna.Testing.RemoteObject), "RemoteObject.rem", WellKnownObjectMode.Singleton);
            try
            {
                ExcelDna.Testing.ExcelStartupEvent.Set();
            }
            catch (System.Threading.WaitHandleCannotBeOpenedException)
            {
                UnregisterChannel();
            }
        }

        public void AutoClose()
        {
            UnregisterChannel();
        }

        public static IpcChannel RegisterIpcChannel(string name, string portName, bool ensureSecurity)
        {
            var ipcChannel = new IpcChannel(
                properties: new Hashtable
                {
                    ["name"] = name,
                    ["portName"] = portName,
                },
                clientSinkProvider: new BinaryClientFormatterSinkProvider { },
                serverSinkProvider: new BinaryServerFormatterSinkProvider { TypeFilterLevel = TypeFilterLevel.Full });

            ChannelServices.RegisterChannel(ipcChannel, ensureSecurity);
            return ipcChannel;
        }

        private void UnregisterChannel()
        {
            if (channel != null)
            {
                ChannelServices.UnregisterChannel(channel);
                channel = null;
            }
        }

        private IpcChannel channel;
    }
}
