using System;
using System.ComponentModel;
using System.Security.Permissions;
using System.Runtime.Serialization;

namespace WecAdmin
{
    /// <summary>
    /// Exception specific to calling Event Collector API commands
    /// </summary>
    [Serializable]
    public class EventCollectorApiException : Exception
    {
        /// <summary>
        /// Initializes a new EventCollectorApiException object that includes the Win32API and Event Collector API name.
        /// </summary>
        /// <param name="Message">Message describing the exception.</param>
        /// <param name="win32Error">Win32 Error code from the native call</param>
        /// <param name="EcApiName">Name of the API called</param>
        public EventCollectorApiException(string Message, int win32Error, string EcApiName) : base(Message)
        {
            EventCollectorApi = EcApiName;
            Win32ErrorCode = win32Error;
            Win32ErrorMessage = new Win32Exception(Win32ErrorCode).Message;
            SubscriptionName = null;
            EventSourceName = null;
        }

        /// <summary>
        /// Initializes a new EventCollectorApiException object that includes the Win32API and Event Collector API name, and Target subscription name.
        /// </summary>
        /// <param name="Message">Message describing the exception.</param>
        /// <param name="win32Error">Win32 Error code from the native call</param>
        /// <param name="EcApiName">Name of the API called</param>
        /// <param name="TargetSubscriptionName">Target Subscription name</param>
        public EventCollectorApiException(string Message, int win32Error, string EcApiName, string TargetSubscriptionName) : base(Message)
        {
            SubscriptionName = TargetSubscriptionName;
            EventCollectorApi = EcApiName;
            Win32ErrorCode = win32Error;
            Win32ErrorMessage = new Win32Exception(Win32ErrorCode).Message;
            EventSourceName = null;
        }

        /// <summary>
        /// Initializes a new EventCollectorApiException object that includes the Win32API and Event Collector API name, Target subscription name, and Event Source name.
        /// </summary>
        /// <param name="Message">Message describing the exception.</param>
        /// <param name="win32Error">Win32 Error code from the native call</param>
        /// <param name="EcApiName">Name of the API called</param>
        /// <param name="TargetSubscriptionName">Target Subscription name</param>
        /// <param name="TargetEventSourceName">Target Event source name referenced in the API call.</param>
        public EventCollectorApiException(string Message, int win32Error, string EcApiName, string TargetSubscriptionName, string TargetEventSourceName) : base(Message)
        {
            SubscriptionName = TargetSubscriptionName;
            EventCollectorApi = EcApiName;
            Win32ErrorCode = win32Error;
            Win32ErrorMessage = new Win32Exception(Win32ErrorCode).Message;
            EventSourceName = TargetEventSourceName;

        }

        [SecurityPermission(SecurityAction.Demand, SerializationFormatter = true)]
        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
            // add other values to object data here
            // example:
            // info.AddValue("CheckedOut", _CheckedOut);
            info.AddValue("EventCollectorApi", this.EventCollectorApi);
            info.AddValue("Win32Errorcode", this.Win32ErrorCode);
            info.AddValue("Win32ErrorMessage", this.Win32ErrorMessage);
        }

        /// <summary>
        /// Event Source name if supplied, null if not supplied.
        /// </summary>
        public string EventSourceName { get; private set; }

        /// <summary>
        /// Name of the Subscription referenced in the call, null if not supplied.
        /// </summary>
        public string SubscriptionName { get; private set; }

        /// <summary>
        /// Name of the Event Collector API called.
        /// </summary>
        public string EventCollectorApi { get; private set; }

        /// <summary>
        /// Win32 error code returned from the underlying API.
        /// </summary>
        public Int32 Win32ErrorCode { get; private set; }

        /// <summary>
        /// Localized message for the supplied win32 error code.
        /// </summary>
        public string Win32ErrorMessage { get; private set; }
    } // public class EventCollectorApiException : Exception

}
