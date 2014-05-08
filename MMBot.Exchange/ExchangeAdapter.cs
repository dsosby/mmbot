
using Common.Logging;

namespace MMBot.Exchange
{
    public class ExchangeAdapter : Adapter
    {
        public ExchangeAdapter(ILog logger, string adapterId)
            : base(logger, adapterId)
        {
        }

        public override System.Threading.Tasks.Task Run()
        {
            throw new System.NotImplementedException();
        }

        public override System.Threading.Tasks.Task Close()
        {
            throw new System.NotImplementedException();
        }
    }
}
