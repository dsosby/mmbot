﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using SystemTask = System.Threading.Tasks;
using System.Timers;
using Common.Logging;
using Microsoft.Exchange.WebServices.Data;

namespace MMBot.Exchange
{
    public class ExchangeAdapter : Adapter
    {
        private string Email { get; set; }
        private string Password { get; set; }
        private string ExchangeUrl { get; set; }
        private int RetrieveCount { get; set; }
        private double PollPeriod { get; set; }
        private int MaxMessagesSaved { get; set; }

        private ExchangeService Service { get; set; }

        private Timer InboxPoll { get; set; }

        private DateTime SinceLastPoll { get; set; }

        private List<EmailMessage> Messages { get; set; }

        public ExchangeAdapter(ILog logger, string adapterId)
            : base(logger, adapterId)
        {
            RetrieveCount = 10;
            PollPeriod = TimeSpan.FromSeconds(10).TotalMilliseconds;
            MaxMessagesSaved = 100;
            Messages = new List<EmailMessage>(MaxMessagesSaved);
        }

        public override void Initialize(Robot robot)
        {
            base.Initialize(robot);
            Configure();

            Service = new ExchangeService
            {
                Credentials = new WebCredentials(Email, Password)
            };

            InitializeExchangeUrl();

            //TODO: Don't use poll, use streaming notifications
            InboxPoll = new Timer(PollPeriod);
            InboxPoll.Elapsed += PollInbox;
            InboxPoll.AutoReset = false;
        }

        private void InitializeExchangeUrl()
        {
            if (string.IsNullOrEmpty(ExchangeUrl))
            {
                Logger.Info("Autodiscovering Exchange service url...");
                Service.AutodiscoverUrl(Email);
            }
            else
            {
                Service.Url = new System.Uri(ExchangeUrl);
            }

            Logger.Info("Exchange service url is " + Service.Url);
        }

        private void Configure()
        {
            Email = Robot.GetConfigVariable("MMBOT_EXCHANGE_EMAIL");
            Password = Robot.GetConfigVariable("MMBOT_EXCHANGE_PASSWORD");
            ExchangeUrl = Robot.GetConfigVariable("MMBOT_EXCHANGE_URL");

            //TODO: Folder, search criteria, retrieve count, poll period
        }

        private void PollInbox(object sender, ElapsedEventArgs e)
        {
            Logger.Debug("Polling inbox");

            var pollSearch = new SearchFilter.IsGreaterThan(ItemSchema.DateTimeReceived, SinceLastPoll);
            var messageList = SearchInbox(pollSearch);
            Logger.Info(string.Format("Processing {0} messages", messageList.Count));

            SaveMessages(messageList);
            foreach (var message in messageList) ProcessMessage(message.Id);

            //Setup for next poll
            //TODO: Is it better to use last message timestamp?
            SinceLastPoll = DateTime.Now;

            Logger.Debug("Scheduling next poll");
            InboxPoll.Start();
        }

        private void ProcessMessage(ItemId id)
        {
            Logger.Info("Getting info for new mail message... ");
            var props = new PropertySet(
                ItemSchema.Id,
                ItemSchema.Subject,
                ItemSchema.UniqueBody,
                ItemSchema.IsFromMe,
                EmailMessageSchema.From
            );
            props.RequestedBodyType = BodyType.Text;
            var message = EmailMessage.Bind(Service, id, props);

            var user = Robot.GetUser(
                message.From.Address,
                message.From.Name,
                message.Id.UniqueId,
                Id);

            var messageBody = message.UniqueBody.Text.Trim();

            Logger.Info(string.Format("Received message from {0}: {1}", user.Id, messageBody));
            if (string.IsNullOrEmpty(messageBody))
            {
                Logger.Info("Skipping empty message: " + message.Subject);
                return;
            }

            var robotMessage = new TextMessage(user, messageBody);
            SystemTask.Task.Run(() => Robot.Receive(robotMessage));
        }

        private void SaveMessages(IEnumerable<EmailMessage> incomingMessages)
        {
            // TODO: Determine when Robot has processed message and remove it from Messages
            // Or can I pass the ID in the Message and then fetch it once I get a callback?

            Messages.AddRange(incomingMessages);
            var removeCount = Messages.Count() - MaxMessagesSaved;
            if (removeCount > 0)
            {
                Logger.Info("Removing " + removeCount + " messages from saved messages");
                Messages.RemoveRange(0, removeCount);
            }
        }

        private ICollection<EmailMessage> SearchInbox(SearchFilter filter)
        {
            var itemView = new ItemView(RetrieveCount);
            var results = Service.FindItems(WellKnownFolderName.Inbox, filter, itemView);
            return results.Items.Select(r => r as EmailMessage).Where(e => e != null).ToList();
        }

        public override async SystemTask.Task Send(Envelope envelope, params string[] messages)
        {
            if (messages == null || !messages.Any()) return;

            var replyToId = envelope.User.Room;
            var replyTo = Messages.FirstOrDefault(m => m.Id.UniqueId == replyToId);

            if (replyTo == null)
            {
                Logger.Info("Could not find parent message for " + replyToId);
                return;
            }

            var response = string.Join("<br>", messages);

            //TODO: Doesn't seem to be replying all
            //TODO: Make these messages prettier to match Outlook styling (HR, use Calibri/sans-serif)
            Logger.Info(string.Format("Replying to {0}: {1}", replyTo.From.Name, response));
            replyTo.Reply(response, replyAll: true);
        }

        public override async SystemTask.Task Run()
        {
            //TODO: Make this actually asynchronous
            SinceLastPoll = DateTime.Now;
            InboxPoll.Start();
        }

        public override async SystemTask.Task Close()
        {
            //TODO: Could wrap this in a task I guess to get rid of await warning
            InboxPoll.Stop();
        }
    }
}
