using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

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
            this.Locale = "en-us";
            this.SubscriptionDescription = string.Empty;
        }

        public bool CreateSubscription()
        {
            throw new NotImplementedException();

            // call validate subscription
            // get subscription handle with create/new
            // call setSubscriptionSettings
            // close subscriptionHandle
        }
        
        public bool ValidateSubscription()
        {
            throw new NotImplementedException();
        }

        public static bool CreateSubscriptionFromXml(XmlDocument SubscriptionConfigurationData)
        {
            // implementation reference: https://msdn.microsoft.com/en-us/library/bb870971(v=vs.85).aspx
            if (null == SubscriptionConfigurationData)
            {
                throw new ArgumentNullException("SubscriptionConfiguation parameter cannot be null");
            }

            // WEC subscriptions have a namespace associated with them, use for querying XML.
            XmlNamespaceManager ecNsMgr = new XmlNamespaceManager(SubscriptionConfigurationData.NameTable);
            ecNsMgr.AddNamespace("ec", @"http://schemas.microsoft.com/2006/03/windows/events/subscription");
            SubscriptionConfiguration.DeliveryConfiguationMode subscriptionDeliveryMode = SubscriptionConfiguration.DeliveryConfiguationMode.Normal;

            XmlNode subscriptionName = SubscriptionConfigurationData.DocumentElement.SelectSingleNode("//ec:Subscription/ec:SubscriptionId", ecNsMgr);
            if (null == subscriptionName || subscriptionName.InnerText == string.Empty)
            {
                throw new ArgumentException("Required configuration node: /Subscription/SubscriptionId missing or empty");

            }
            // TODO:PARAMETERCHECK - make sure the subscription ID passes the filesystem + registry + eventlog name requirements.

            XmlNode subscriptionType = SubscriptionConfigurationData.DocumentElement.SelectSingleNode("//ec:Subscription/ec:SubscriptionType", ecNsMgr);
            if (null == subscriptionType || (subscriptionType.InnerText != "CollectorInitiated" && subscriptionType.InnerText != "SourceInitiated"))
            {
                throw new ArgumentException("Required configuration node: /Subscription/SubscriptionType missing or invalid value. Valid values: [SourceInitiated, CollectorInitiated]");
            }

            XmlNode enabled = SubscriptionConfigurationData.DocumentElement.SelectSingleNode("//ec:Subscription/ec:Enabled", ecNsMgr);
            if (null == enabled || (enabled.InnerText != "true" && enabled.InnerText != "false"))
            {
                throw new ArgumentException("Required configuration node: /Subscription/Enabled missing or invalid value. Valid values: [true, false]");
            }


            XmlNode subscriptionUri = SubscriptionConfigurationData.DocumentElement.SelectSingleNode("//ec:Subscription/ec:Uri", ecNsMgr);
            if (null == enabled || (subscriptionUri.InnerText != "http://schemas.microsoft.com/wbem/wsman/1/windows/EventLog"))
            {
                throw new ArgumentException("Required configuration node: /Subscription/Uri missing or invalid value. Valid value(s): [http://schemas.microsoft.com/wbem/wsman/1/windows/EventLog]");
            }

            XmlNode configurationMode = SubscriptionConfigurationData.DocumentElement.SelectSingleNode("//ec:Subscription/ec:ConfigurationMode", ecNsMgr);
            if (null == enabled)
            {
                throw new ArgumentException("Required configuration node: /Subscription/Uri node missing");
            }
            else
            {
                // check value against 
                switch (configurationMode.InnerText.ToLower())
                {
                    case "normal":
                        subscriptionDeliveryMode = SubscriptionConfiguration.DeliveryConfiguationMode.Normal;
                        break;
                    case "minbandwidth":
                        subscriptionDeliveryMode = SubscriptionConfiguration.DeliveryConfiguationMode.MinBandwidth;
                        break;
                    case "minlatency":
                        subscriptionDeliveryMode = SubscriptionConfiguration.DeliveryConfiguationMode.MinLatency;
                        break;
                    case "custom":
                        subscriptionDeliveryMode = SubscriptionConfiguration.DeliveryConfiguationMode.Custom;
                        break;
                    default:
                        throw new ArgumentException("ConfigurationMode valud is invalid. Valid values: [Normal, MinBandwidth, MinLatency, Custom]");
                }
            }


            XmlNode description = SubscriptionConfigurationData.DocumentElement.SelectSingleNode("//ec:Subscription/ec:Description", ecNsMgr);
            if (null == description)
            {
                //TODO:DECIDE -- really needed or just a nice to have?
                throw new ArgumentException("Required configuration node: /Subscription/Description missing");
            }


            // Required configuration items
            /*
             * Xml Path : Name
             * /Subscription/SubscriptionId : subscription name
             * /Subscription/SubscrpitionType : sub type [SourceInitiated || CollectorInitiated]
               /Subscription/Enabled
               /Subscription/Uri  (MUST be: http://schemas.microsoft.com/wbem/wsman/1/windows/EventLog)
             * /Subscription/ConfigurationMode [Normal, Custom, MinLatency, MinBandwitdh]
             * 
             * */

            // check settings for required property values

            //* /Subscription/Description
            // check existing subscriptions for name collision


            // create handle to subscription
            //      EC_READ_ACCESS | EC_WRITE_ACCESS, 
            //      EC_CREATE_NEW

            // for delivery options - if "custom" then check for batch/delay values being set
            // set ALL the properties
            // save
            // example/ref: https://msdn.microsoft.com/en-us/library/bb870971(v=vs.85).aspx

            throw new NotImplementedException("CreateSubscription - implement me fully!");

        } // public static bool CreateSubscriptionFromXml(XmlDocument SubscriptionConfiguration)


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
