using Xunit.Abstractions;
using Xunit.Sdk;

namespace ExcelDna.Testing.Remote
{
    internal class MessageBusProxy : IMessageBus
    {
        public delegate bool BusMessageDelegate(IMessageSinkMessage message);

        public MessageBusProxy(BusMessageDelegate proxy)
        {
            this.proxy = proxy;
        }

        public void Dispose()
        {
        }

        public bool QueueMessage(IMessageSinkMessage message)
        {
            return proxy(message);
        }

        private BusMessageDelegate proxy;
    }
}
