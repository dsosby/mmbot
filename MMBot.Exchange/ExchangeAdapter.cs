using System;
using System.Collections.Generic;
using System.Linq;
using System.Timers;
using Common.Logging;
using Microsoft.Exchange.WebServices.Data;

namespace MMBot.Exchange
{
    public class ExchangeAdapter : Adapter
    {
        private string Email { get; set; }
        private string Password { get; set; }
        private int RetrieveCount { get; set; }

        private ExchangeService Service { get; set; }

        private Timer InboxPoll { get; set; }

        private readonly SearchFilter _unreadFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And,
                new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));

        private List<EmailMessage> MessageQueue = new List<EmailMessage>();

        public ExchangeAdapter(ILog logger, string adapterId)
            : base(logger, adapterId)
        {
            RetrieveCount = 10;
        }

        public override void Initialize(Robot robot)
        {
            base.Initialize(robot);
            Configure();

            Service = new ExchangeService
            {
                Credentials = new WebCredentials(Email, Password)
            };

            InboxPoll = new Timer(TimeSpan.FromSeconds(10).TotalMilliseconds);
            InboxPoll.Elapsed += PollInbox;
            InboxPoll.AutoReset = true;

        }

        private void Configure()
        {
            Email = Robot.GetConfigVariable("MMBOT_EXCHANGE_EMAIL");
            Password = Robot.GetConfigVariable("MMBOT_EXCHANGE_PASSWORD");
        }

        private void PollInbox(object sender, ElapsedEventArgs e)
        {
            Logger.Debug("Polling Inbox");
            var unread = GetUnread();
            MessageQueue = MessageQueue.Union(unread).Distinct().ToList();
        }

        public ICollection<EmailMessage> GetUnread()
        {
            var unread = new ItemView(RetrieveCount);
            var results = Service.FindItems(WellKnownFolderName.Inbox, _unreadFilter, unread);
            return results.Items.Select(r => r as EmailMessage).Where(e => e != null).ToList();
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
