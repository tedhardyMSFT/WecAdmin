using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WecAdmin
{
    class SubscriptionConfiguration
    {
        public enum DeliveryConfiguationMode
        {
            /// <summary>
            /// This option ensures reliable delivery of events and does not attempt to conserve bandwidth. It is the appropriate choice unless you need tighter control over bandwidth usage or need forwarded events delivered as quickly as possible. It uses pull delivery mode, batches 5 items at a time and sets a batch timeout of 15 minutes
            /// </summary>
            Normal = 0,
            /// <summary>
            /// This option ensures that the use of network bandwidth for event delivery is strictly controlled. It is an appropriate choice if you want to limit the frequency of network connections made to deliver events. It uses push delivery mode and sets a batch timeout of 6 hours. In addition, it uses a heartbeat interval of 6 hours.
            /// </summary>
            MinBandwidth = 1,
            MinLatency = 2,
            Custom = 3
        }

        public enum ContentFormat
        {

            RenderedText = 0,

            Events = 1
        }

        public string SubscriptionId { get; set; }

        public string SubscriptionDescription { get; set; }

        public string DestinationChannel { get; set; }

        public string EventFilterQuery { get; set; }

        public string WsManUri { get; set; }

        public bool Enabled { get; set; }

        public WecAdmin.SubscriptionConfiguration.DeliveryConfiguationMode ConfigurationMode { get; set; }

        public int BatchingMaxItems { get; set; }

        public int BatchingMaxLatencyTime { get; set; }

        public int HeartbeatInterval { get; set; }

        public DateTime SubscriptionExpires { get; set; }

        public bool ReadExistingEvents { get; set; }

        public WecAdmin.SubscriptionConfiguration.ContentFormat EventContentformat { get; set; }

        public string Locale { get; set; }

        public string AllowedSourceNonDomainComputers { get; set; }

        public string AllowedSourceDomainComputers { get; set; }

        private IntPtr ecSubscriptionHandle = IntPtr.Zero;

        /// <summary>
        /// Initializes an empty instance of the Subscription Configuration object
        /// </summary>
        public SubscriptionConfiguration()
        {
            // set default values
            this.WsManUri = "http://schemas.microsoft.com/wbem/wsman/1/windows/EventLog";
            this.SubscriptionDescription = string.Empty;
        }

        public bool CreateSubscription()
        {
            throw new NotImplementedException();

            // get subscription handle with create/new
            // call setSubscriptionSettings
            // close subscriptionHandle
        }

        public bool UpdateSubscriptionSettings()
        {
            throw new NotImplementedException();

            // get subscription handle with open/write
            // call setSubscriptionSettings
            // close subscriptionHandle

        }

        private int setSubscriptionSettings()
        {
            throw new NotImplementedException();
        }
    }
}
