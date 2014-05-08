using System.Timers;
using Common.Logging;
using Microsoft.Exchange.WebServices.Data;

namespace MMBot.Exchange
{
    public class ExchangeAdapter : Adapter
    {
        private string Email { get; set; }
        private string Password { get; set; }

        private ExchangeService Service { get; set; }

        private Timer InboxPoll { get; set; }

        private readonly SearchFilter _unreadFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And,
                new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));

        public ExchangeAdapter(ILog logger, string adapterId)
            : base(logger, adapterId)
        {
        }

        private void Configure()
        {
            Email = Robot.GetConfigVariable("MMBOT_EXCHANGE_EMAIL");
            Password = Robot.GetConfigVariable("MMBOT_EXCHANGE_PASSWORD");
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
