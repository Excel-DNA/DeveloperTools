using Xunit.Abstractions;
using Xunit.Sdk;

namespace ExcelDna.Testing.Remote
{
    internal class MessageBusProxy : IMessageBus
    {
        public delegate void BusMessageDelegate(IMessageSinkMessage message);

        public MessageBusProxy(BusMessageDelegate proxy)
        {
            this.proxy = proxy;
        }

        public void Dispose()
        {
        }

        public bool QueueMessage(IMessageSinkMessage message)
        {
            try
            {
                proxy(message);
            }
            catch
            {
            }

            return true;
        }

        private BusMessageDelegate proxy;
    }
}
