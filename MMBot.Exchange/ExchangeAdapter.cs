using System;
using System.Collections.Generic;
using System.Linq;
using SystemTask = System.Threading.Tasks;
using Common.Logging;
using Microsoft.Exchange.WebServices.Data;

namespace MMBot.Exchange
{
    public class ExchangeAdapter : Adapter
    {
        private string Email { get; set; }
        private string Password { get; set; }
        private string ExchangeUrl { get; set; }

        private PropertySet EmailProperties { get; set; }

        private ExchangeService Service { get; set; }

        private StreamingSubscriptionConnection ExchangeConnection { get; set; }

        private bool IsRunning { get; set; }

        private List<EmailMessage> Messages { get; set; }


        public ExchangeAdapter(ILog logger, string adapterId)
            : base(logger, adapterId)
        {
            Messages = new List<EmailMessage>();

            EmailProperties = new PropertySet(
                ItemSchema.Id,
                ItemSchema.Subject,
                ItemSchema.UniqueBody,
                ItemSchema.IsFromMe,
                ItemSchema.DateTimeReceived,
                EmailMessageSchema.From
            );
            EmailProperties.RequestedBodyType = BodyType.Text;
        }

        public override void Initialize(Robot robot)
        {
            base.Initialize(robot);
            Configure();

            if (string.IsNullOrEmpty(Email) ||
                string.IsNullOrEmpty(Password))
            {
                Logger.Warn("Exchange Adapter requires MMBOT_EXCHANGE_EMAIL and MMBOT_EXCHANGE_PASSWORD");
                return;
            }

            Service = new ExchangeService
            {
                Credentials = new WebCredentials(Email, Password)
            };

            InitializeExchangeUrl();

            var newMailSubscription = Service.SubscribeToStreamingNotifications(
                new FolderId[] {WellKnownFolderName.Inbox},
                EventType.NewMail);

            ExchangeConnection = new StreamingSubscriptionConnection(Service, 30);
            ExchangeConnection.AddSubscription(newMailSubscription);
            ExchangeConnection.OnNotificationEvent += OnExchangeNotification;
            ExchangeConnection.OnDisconnect += OnExchangeDisconnect;
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

            //TODO: Folder? Subject filter? From domain filter? Subscription timeout?
        }

        private void OnExchangeDisconnect(object sender, SubscriptionErrorEventArgs args)
        {
            bool isRecoverable = args.Exception == null;

            if (IsRunning && isRecoverable)
            {
                Logger.Info("Restarting Exchange subscription");
                ExchangeConnection.Open();
            }
            else
            {
                Logger.Info("Exchange service disconnected: " + IsRunning + " " + args.Exception);
            }
        }

        private void OnExchangeNotification(object sender, NotificationEventArgs args)
        {
            var service = args.Subscription.Service;
            var emailIds = args.Events
                .OfType<ItemEvent>()
                .Select(i => i.ItemId);
            var emails = service.BindToItems(emailIds, EmailProperties)
                .Select(r => r.Item as EmailMessage)
                .ToList();

            SaveMessages(emails);
            foreach (var email in emails) ProcessMessage(email);
        }

        private void ProcessMessage(EmailMessage message)
        {
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
            //Only save messages for an hour or so
            Messages.AddRange(incomingMessages);
            Messages = Messages
                .Where(m => m.DateTimeReceived > DateTime.Now.AddHours(-1))
                .ToList();
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
            if (Service == null) return;

            IsRunning = true;
            ExchangeConnection.Open();
        }

        public override async SystemTask.Task Close()
        {
            if (Service == null) return;

            IsRunning = false;
            ExchangeConnection.Close();
        }
    }
}
