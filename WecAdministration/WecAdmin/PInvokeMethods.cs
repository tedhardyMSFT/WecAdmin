﻿//
// function and parameter descriptions are directly from MSDN whereever possible.
//
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace WecAdmin
{
    /// <summary>
    /// The EC_VARIANT structure contains event collector data (subscription data) or property values.
    /// </summary>
    [StructLayout(LayoutKind.Explicit)]
    public struct EC_VARIANT
    {
        /// <summary>
        /// A boolean value in a byte.
        /// </summary>
        [FieldOffset(0)]
        internal byte BooleanVal;
        /// <summary>
        ///  an unsigned 32-bit integer value
        /// </summary>
        [FieldOffset(0)]
        internal IntPtr UInt32Val;
        /// <summary>
        /// A pointer to a ULONGLONG value
        /// </summary>
        [FieldOffset(0)]
        internal IntPtr DateTimeVal;
        /// <summary>
        /// A pointer to a null-terminated Unicode value
        /// </summary>
        [FieldOffset(0)]
        internal IntPtr StringValue;
        /// <summary>
        /// A pointer to a hexadecimal binary value
        /// </summary>
        [FieldOffset(0)]
        internal IntPtr BinaryVal;
        /// <summary>
        /// A pointer to an array of Boolean values
        /// </summary>
        [FieldOffset(0)]
        internal IntPtr BooleanArr;
        /// <summary>
        /// A pointer to an arrya of signed 32-bit integer values
        /// </summary>
        [FieldOffset(0)]
        internal IntPtr Int32Arr;
        /// <summary>
        /// A pointer to an array of null-terminated strings
        /// </summary>
        [FieldOffset(0)]
        internal IntPtr StringArr;
        /// <summary>
        /// The number of elements (not legnth) in bytes. Used for arrays and binary or string types.
        /// </summary>
        [FieldOffset(8)]
        internal UInt32 Count;
        /// <summary>
        /// The type of data in the structure. Use a value from the EC_VARIANT_TYPE enumeration to specify the type. When the type is specified, you can use any of the union members to access the actual value. For example, if the type is EcVarTypeDateTime, then the value is DateTimeVal in the EC_VARIANT structure.
        /// </summary>
        [FieldOffset(12)]
        internal UInt32 Type;
    } // public struct EC_VARIANT

    /// <summary>
    /// Defines the P/Invoke methods and enumerations for accessing Event Collector
    /// </summary>
    class NativeMethods
    {
        #region Enumeration Definitions

        /// <summary>
        /// The EC_SUBSCRIPTION_CREDENTIALS_TYPE enumeration specifies the type of credentials to use when communicating with event sources.
        /// </summary>
        public enum EC_SUBSCRIPTION_CREDENTIALS_TYPE
        {
            /// <summary>
            /// Negotiate with event sources to specify a proper authentication type without specifying a username and password for the subscription credentials.
            /// </summary>
            EcSubscriptionCredDefault = 0,
            /// <summary>
            /// WinRM will negotiate with event sources to specify a proper authentication type for the subscription credentials.
            /// </summary>
            EcSubscriptionCredNegotiate = 1,
            /// <summary>
            /// Use digest authentication for the subscription credentials.
            /// </summary>
            EcSubscriptionCredDigest = 2,
            /// <summary>
            /// Send a username and password to use as credentials for the subscription.
            /// </summary>
            EcSubscriptionCredBasic = 3,
            /// <summary>
            /// Use the local computer's domain account credentials to create a subscription instead of using user credentials for the subscription. This has the advantage of not having to manage user accounts and password expiration to simplify long lasting subscription management.
            /// </summary>
            EcSubscriptionCredLocalMachine = 4

        } // public enum EC_SUBSCRIPTION_CREDENTIALS_TYPE

        /// <summary>
        /// The EC_SUBSCRIPTION_CONTENT_FORMAT enumeration specifies how events will be rendered on the computer that sends the events before the events are sent to the event collector computer
        /// </summary>
        public enum EC_SUBSCRIPTION_CONTENT_FORMAT
        {
            /// <summary>
            /// When an event is received, the Event Collector service sends an event as the received event to an event log. The service sends the raw event data only, and not any localized event data.
            /// </summary>
            EcContentFormatEvents = 1,
            /// <summary>
            /// When an event is received, the Event Collector service sends an event as rendered text to an event log. The service sends raw event data and localized event information.
            /// </summary>
            EcContentFormatRenderedText = 2
        } // public enum EcContentFormatRenderedText

        /// <summary>
        /// Specifies a value that identifies a property of the runtime status of an event source or a subscription.
        /// </summary>
        public enum EC_SUBSCRIPTION_RUNTIME_STATUS_INFO_ID
        {
            /// <summary>
            /// Get the status of an active or inactive subscription or an event source. This will return an unsigned 32-bit integer value from the EC_SUBSCRIPTION_RUNTIME_STATUS_ACTIVE_STATUS enumeration.
            /// </summary>
            EcSubscriptionRunTimeStatusActive = 0,
            /// <summary>
            /// Get the last error status of a subscription or an event source. This will return an EcVarTypeUInt32 value.
            /// </summary>
            EcSubscriptionRunTimeStatusLastError = 1,
            /// <summary>
            /// Get the last error message for a subscription or an event source. This will return an EcVarTypeString value.
            /// </summary>
            EcSubscriptionRunTimeStatusLastErrorMessage = 2,
            /// <summary>
            /// Get the time that the last error occurred for a subscription or an event source. This will return an EcVarTypeDateTime value.
            /// </summary>
            EcSubscriptionRunTimeStatusLastErrorTime = 3,
            /// <summary>
            /// Get the next time that the subscription or an event source will try to run (after an error). This will return an EcVarTypeDateTime value.
            /// </summary>
            EcSubscriptionRunTimeStatusNextRetryTime = 4,
            /// <summary>
            /// Get the event sources for the subscription. For collector initiated subscriptions, 
            /// this list will be identical to the one in the subscription's configuration. For source initiated subscriptions, 
            /// this list will be the set of event sources that collector has heard from in the last 30 days. 
            /// This list is persistent across reboots of the event collector. This will return an EcVarTypeString value.
            /// </summary>
            EcSubscriptionRunTimeStatusEventSources = 5,
            /// <summary>
            /// Get the last time that a heartbeat (a signal used to signify the subscription is working) occurred 
            /// for a subscription or an event source. This will return an EcVarTypeDateTime value.
            /// </summary>
            EcSubscriptionRunTimeStatusLastHeartbeatTime = 6
        } // public enum EC_SUBSCRIPTION_RUNTIME_STATUS_INFO_ID

        /// <summary>
        /// Specifies different configuration modes that change the default settings for a subscription.
        /// </summary>
        public enum EC_SUBSCRIPTION_CONFIGURATION_MODE
        {
            /// <summary>
            /// This mode is used when an administrator needs the events to be delivered reliably and for the subscription to work with minimal configuration, and when network usage is not a concern.
            /// This mode sets the default subscription delivery mode to pull subscriptions.
            /// </summary>
            EcConfigurationModeNormal = 0,
            /// <summary>
            /// This subscription mode allows custom values for the DeliveryMode property, the DeliveryMaxItems property, the DeliveryMaxLatencyTime, and the HeartBeatInterval property.
            /// </summary>
            EcConfigurationModeCustom = 1,
            /// <summary>
            /// This mode is used for alerts and critical events because it configures the subscription to send events as soon as they occur with minimal delay.
            /// This mode sets the default subscription delivery mode to push subscriptions.
            /// </summary>
            EcConfigurationModeMinLatency = 2,
            /// <summary>
            /// This mode is used when network activity is controllable, and when network usage is expensive.
            /// This mode sets the default subscription delivery mode to push subscriptions.
            /// </summary>
            EcConfigurationModeMinBandwidth = 3
        } // enum EC_SUBSCRIPTION_CONFIGURATION_MODE

        /// <summary>
        ///  	Specifies how events are delivered through an event subscription (using a push or pull model).
        /// </summary>
        public enum EC_SUBSCRIPTION_DELIVERY_MODE
        {
            /// <summary>
            /// Events are delivered through the subscription using the pull model.
            /// </summary>
            EcDeliveryModePull = 1,
            /// <summary>
            /// Events are delivered through the subscription using the push model.
            /// </summary>
            EcDeliveryModePush = 2
        } // enum EC_SUBSCRIPTION_DELIVERY_MODE


        /// <summary>
        /// Defines values to identify event subscription properties used for subscription configuration.
        /// </summary>
        public enum EC_SUBSCRIPTION_PROPERTY_ID
        {
            /// <summary>
            /// The Enabled property of the subscription that is used to enable or disable the subscription or
            /// obtain the current status of a subscription.
            /// This property is an EcVarTypeBoolean value.
            /// </summary>
            EcSubscriptionEnabled = 0,
            /// <summary>
            /// The EventSources property of the subscription that contains a collection of information about
            /// the local or remote computers (event sources) that can forward events to the event collector. 
            /// This property is a handle to an array (an EcVarObjectArrayPropertyHandle value). 
            /// This value is typically used for collector initiated subscriptions. 
            /// It can be used for source initiated subscriptions to disable the collection of events from
            /// a particular event source.
            /// </summary>
            EcSubscriptionEventSources = 1,
            /// <summary>
            /// The EventSourceAddress property of the subscription that contains the IP address
            /// or fully qualified domain name (FQDN) of the local or remote computer (event source) from
            /// which the events are collected. 
            /// This property is an EcVarTypeString value.
            /// </summary>
            EcSubscriptionEventSourceAddress = 2,
            /// <summary>
            /// The EventSourceEnabled property of the subscription that is used to enable or disable an event source. 
            /// This property is an EcVarTypeBoolean value.
            /// </summary>
            EcSubscriptionEventSourceEnabled = 3,
            /// <summary>
            /// The EventSourceUserName property of the subscription that contains the user name, 
            /// which is used by the remote computer (event source) to authenticate the user. 
            /// This property is an EcVarTypeString value. 
            /// This property cannot be used for source initiated subscriptions.
            /// </summary>
            EcSubscriptionEventSourceUserName = 4,
            /// <summary>
            /// The EventSourcePassword property of the subscription that contains the password, 
            /// which is used by the remote computer (event source) to authenticate the user. 
            /// This property is an EcVarTypeString value. 
            /// This property cannot be used for source initiated subscriptions.
            /// </summary>
            EcSubscriptionEventSourcePassword = 5,
            /// <summary>
            /// The Description property of the subscription that contains a description of the subscription. 
            /// This property is an EcVarTypeString value.
            /// </summary>
            EcSubscriptionDescription = 6,
            /// <summary>
            /// The URI property of the subscription that contains the URI, which is used by WS-Management to 
            /// connect to a computer. For example, the URI can be http://schemas.microsoft.com/wbem/wsman/1/logrecord/sel 
            /// for hardware events or it can be http://schemas.microsoft.com/wbem/wsman/1/windows/EventLog for events
            /// that are published in the event log. 
            /// This property is an EcVarTypeString value.
            /// </summary>
            EcSubscriptionURI = 7,
            /// <summary>
            /// The ConfigurationMode property of the subscription that specifies how events are delivered to the subscription. 
            /// This property is an EcVarTypeUInt32 value from the EC_SUBSCRIPTION_CONFIGURATION_MODE enumeration.
            /// </summary>
            EcSubscriptionConfigurationMode = 8,
            /// <summary>
            /// The Expires property of the subscription that contains the date when the subscription will end. 
            /// The maximum date that can be used is 3000-12-31T23:59:59.999Z.
            /// If this property is not defined, the subscription will not expire. 
            /// This property is an EcVarTypeDateTime value.
            /// </summary>
            EcSubscriptionExpires = 9,
            /// <summary>
            /// The Query property of the subscription that contains the query, which is used by the event
            /// source for selecting events to forward to the event collector.
            /// This property is an EcVarTypeString value.
            /// </summary>
            EcSubscriptionQuery = 10,
            /// <summary>
            /// The TransportName property of the subscription that specifies the type of transport,
            /// which is used to connect to the remote computer (event source). 
            /// This value can be either HTTP, which is the default, or it can be HTTPS.
            /// This property is an EcVarTypeString value.
            /// </summary>
            EcSubscriptionTransportName = 11,
            /// <summary>
            /// The TransportPort property of the subscription that specifies the port number, which the
            /// transport uses to connect to the remote computer (event source). The default port number
            /// for HTTP is 80 and the default port number for HTTPS is 443.
            /// This property is an EcVarTypeUInt32 value.
            /// </summary>
            EcSubscriptionTransportPort = 12,
            /// <summary>
            /// The DeliveryMode property of the subscription that specifies whether events are delivered
            /// to the subscription with either a push or pull model. 
            /// This property is an EC_SUBSCRIPTION_DELIVERY_MODE enumeration value.
            /// This property cannot be used for source initiated subscriptions.
            /// </summary>
            EcSubscriptionDeliveryMode = 13,
            /// <summary>
            /// The DeliveryMaxItems property of the subscription that specifies the maximum number of
            /// events that can be batched when forwarded from the event sources. 
            /// When the EcSubscriptionDeliveryMode property is set to EcDeliveryModePush, 
            /// this property determines the number of events that are included in a batch sent from the event source. 
            /// When the EcSubscriptionDeliveryMode property is set to EcDeliveryModePull, 
            /// this property determines the maximum number of items that will forwarded from an event source for each request.
            /// This property is an EcVarTypeUInt32 value.
            /// </summary>
            EcSubscriptionDeliveryMaxItems = 14,
            /// <summary>
            /// The DeliveryMaxLatencyTime property of the subscription that specifies how long, in milliseconds,
            /// the event source should wait before sending events (even if it did not collect enough events
            /// to reach the maximum number of items). 
            /// This value is used when the EcSubscriptionDeliveryMode property is set to EcDeliveryModePush.
            /// This property is an EcVarTypeUInt32 value.
            /// </summary>
            EcSubscriptionDeliveryMaxLatencyTime = 15,
            /// <summary>
            /// The HeartbeatInterval property of the subscription that defines the heartbeat time interval,
            /// in milliseconds, which is observed between the sent heartbeat messages.
            /// When the EcSubscriptionDeliveryMode property is set to EcDeliveryModePush, the event collector
            /// uses this property to determine the availability of the event source.
            /// When the EcSubscriptionDeliveryMode property is set to EcDeliveryModePull, the event collector
            /// uses this property to determine the interval between queries to the event source.
            /// This property is an EcVarTypeUInt32 value.
            /// </summary>
            EcSubscriptionHeartbeatInterval = 16,
            /// <summary>
            /// The Locale property of the subscription that specifies the locale (for example, en-us) of the events.
            /// This property is an EcVarTypeString value.
            /// </summary>
            EcSubscriptionLocale = 17,
            /// <summary>
            /// The ContentFormat property of the subscription that specifies the format in which the event content
            /// should be delivered.
            /// This property is an EC_SUBSCRIPTION_CONTENT_FORMAT enumeration value.
            /// </summary>
            EcSubscriptionContentFormat = 18,
            /// <summary>
            /// The LogFile property of the subscription that specifies the log file where the events collected
            /// from the event sources will be stored.
            /// This property is an EcVarTypeString value.
            /// </summary>
            EcSubscriptionLogFile = 19,
            /// <summary>
            /// The PublisherName property of the subscription that contains the name of publisher that the
            /// event collector computer will raise events to the local log as.
            /// This is used when you want to collect events in a log other than the ForwardedEvents log.
            /// This property is an EcVarTypeString value.
            /// </summary>
            EcSubscriptionPublisherName = 20,
            /// <summary>
            /// The CredentialsType property of the subscription that specifies the type of credentials
            /// used in the event subscription.
            /// This property is an EC_SUBSCRIPTION_CREDENTIALS_TYPE enumeration value.
            /// This property cannot be used for source initiated subscriptions.
            /// </summary>
            EcSubscriptionCredentialsType = 21,
            /// <summary>
            /// The CommonUserName property of the subscription that contains the common user name,
            /// which is used by the local and remote computers to authenticate the user.
            /// This property is an EcVarTypeString value.
            /// This property cannot be used for source initiated subscriptions.
            /// </summary>
            EcSubscriptionCommonUserName = 22,
            /// <summary>
            /// The CommonPassword property of the subscription that contains the common password,
            /// which is used by the local and remote computers to authenticate the user.
            /// This property is an EcVarTypeString value.
            /// This property cannot be used for source initiated subscriptions.
            /// </summary>
            EcSubscriptionCommonPassword = 23,
            /// <summary>
            /// The HostName property of the subscription that specifies the fully qualified domain name (FQDN) of the local computer.
            /// This property is used by an event source to forward events and is used in scenarios that
            /// involve multihomed servers that may have multiple FQDNs.
            /// This property is an EcVarTypeString value and must only be used for a push subscription.
            /// </summary>
            EcSubscriptionHostName = 24,
            /// <summary>
            /// The ReadExistingEvents property of the subscription that determines whether to collect
            /// existing events or not.
            /// This property is an EcVarTypeBoolean value.
            /// </summary>
            EcSubscriptionReadExistingEvents = 25,
            /// <summary>
            /// The Dialect property of the subscription that specifies the dialect of the query string.
            /// For example, the dialect for SQL based filters would be SQL, and
            /// the dialect for WMI based filters would be WQL.
            /// This property is an EcVarTypeString value.
            /// </summary>
            EcSubscriptionDialect = 26,
            /// <summary>
            /// The Type property of the subscription that defines whether the subscription is
            /// initiated by an event source or collector.
            /// This property is a EC_SUBSCRPTION_TYPE value.
            /// </summary>
            EcSubscriptionType = 27,
            /// <summary>
            /// The AllowedIssuerCAs property of the subscription that contains the certificate authorities
            /// (CAs) allowed if the subscription uses certificate-based authentication.
            /// This is used for source initiated subscriptions.
            /// This property is an EcVarTypeString value.
            /// </summary>
            EcSubscriptionAllowedIssuerCAs = 28,
            /// <summary>
            /// The AllowedSubjects property of the subscription that contains the subjects that are allowed for
            /// the subscription.
            /// This is used for source initiated subscriptions.
            /// The subject specifies names, such as domain names, for all the event source computers that are allowed
            /// in the subscription.
            /// This property is an EcVarTypeString value.
            /// </summary>
            EcSubscriptionAllowedSubjects = 29,
            /// <summary>
            /// The DeniedSubjects property of the subscription that contains the subjects that are not allowed
            /// for the subscription.
            /// This is used for source initiated subscriptions.
            /// The subject specifies names, such as domain names, for all the event source computers that are
            /// not allowed in the subscription.
            /// This property is an EcVarTypeString value.
            /// </summary>
            EcSubscriptionDeniedSubjects = 30,
            /// <summary>
            /// The AllowedSourceDomainComputers property of the subscription that contains the source computers
            /// that are allowed to send events to the collector computer defined by an SDDL string.
            /// This property is an EcVarTypeString value.
            /// </summary>
            EcSubscriptionAllowedSourceDomainComputers = 31
        } // enum EC_SUBSCRIPTION_PROPERTY_ID

        /// <summary>
        /// Specifies the status of a subscription or an event source with respect to a subscription.
        /// </summary>
        public enum EC_SUBSCRIPTION_RUNTIME_STATUS_ACTIVE_STATUS
        {
            /// <summary>
            /// The subscription or event source is disabled.
            /// </summary>
            EcRuntimeStatusActiveStatusDisabled = 1,
            /// <summary>
            /// The subscription or event source is running.
            /// </summary>
            EcRuntimeStatusActiveStatusActive = 2,
            /// <summary>
            /// The subscription or event source is inactive. You can query the System event log to see the error events sent by the Event Collector service. Use the EC_SUBSCRIPTION_RUNTIME_STATUS_INFO_ID values to obtain information on why the subscription or source is inactive.
            /// </summary>
            EcRuntimeStatusActiveStatusInactive = 3,
            /// <summary>
            /// The subscription or event source is trying to connect for the first time or is retrying after a problem.
            /// When an active subscription runs into a problem, it will retry several times.
            /// </summary>
            EcRuntimeStatusActiveStatusTrying = 4
        } // enum EC_SUBSCRIPTION_RUNTIME_STATUS_ACTIVE_STATUS

        /// <summary>
        /// Defines the values that specify the data types that are used in the Windows Event Collector functions.
        /// </summary>
        public enum EC_VARIANT_TYPE
        {
            /// <summary>
            /// Null content that implies that the element that contains the content does not exist.
            /// </summary>
            EcVarTypeNull = 0,
            /// <summary>
            /// A Boolean value.
            /// </summary>
            EcVarTypeBoolean = 1,
            /// <summary>
            /// An unsigned 32-bit value.
            /// </summary>
            EcVarTypeUInt32 = 2,
            /// <summary>
            /// A ULONGLONG value.
            /// </summary>
            EcVarTypeDateTime = 3,
            /// <summary>
            /// A null-terminated Unicode string
            /// </summary>
            EcVarTypeString = 4,
            /// <summary>
            /// An EC_OBJECT_ARRAY_PROPERTY_HANDLE value.
            /// </summary>
            EcVarObjectArrayPropertyHandle = 5
        } // enum EC_VARIANT_TYPE
        #endregion


        // constants from: https://msdn.microsoft.com/en-us/library/aa385773(v=vs.85).aspx
        /// <summary>
        /// Read access control permission that allows information to be read from the event collector.
        /// </summary>
        public const Int32 EC_READ_ACCESS = 1;
        /// <summary>
        /// Write access control permission that allows information to be written to the event collector.
        /// </summary>
        public const Int32 EC_WRITE_ACCESS = 2;
        /// <summary>
        /// Opens an existing subscription or creates the subscription if it does not exist. Used by the EcOpenSubscription method.
        /// </summary>
        public const Int32 EC_OPEN_ALWAYS = 0;
        /// <summary>
        /// A flag passed to the EcOpenSubscription function specifying that a new subscription should be created.
        /// </summary>
        public const Int32 EC_CREATE_NEW = 1;
        /// <summary>
        /// A flag passed to the EcOpenSubscription function specifying that an existing subscription should be opened.
        /// </summary>
        public const Int32 EC_OPEN_EXISTING = 2;

        /// <summary>
        /// Used to mask out the array bit from the Type property of an EC_VARIANT to extract the type of the variant value.
        /// </summary>
        public const Int32 EC_VARIANT_TYPE_MASK = 0x7f;

        /// <summary>
        /// When this bit is set in the Type property of an EC_VARIANT, the variant contains a pointer to an array of values, rather than the value itself.
        /// </summary>
        public const Int32 EC_VARIANT_TYPE_ARRAY = 0x80;


        /// <summary>
        /// The EcClose function closes a handle received from other Event Collector functions. Any handle returned by an event collector management API call must be closed using this call when the user is finished with the handle. The handle becomes invalid when this function is successfully called.
        /// </summary>
        /// <param name="Object">A valid open handle returned from an event collector management API call.</param>
        /// <returns>True if the function was successful. False if the function failed. Use the GetLastError function to obtain the error code.</returns>
        [DllImport("wecapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool EcClose(
            IntPtr Object
            );


        /// <summary>
        /// The EcDeleteSubscription function deletes an existing subscription that is specified by the SubscriptionName parameter. The function fails if the security descriptor of the subscription does not permit delete access for the calling process. If the subscription is active at the moment this API is called, then the subscription is deactivated.
        /// </summary>
        /// <param name="SubscriptionName">The subscription to be deleted.</param>
        /// <param name="Flags">Reserved, must be 0.</param>
        /// <returns>TRUE The function was successful. FALSE The function failed.Use the GetLastError function to obtain the error code.</returns>
        [DllImport("wecapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool EcDeleteSubscription(
            IntPtr SubscriptionName,
            Int32 Flags
            );

        /// <summary>
        /// The EcEnumNextSubscription function continues the enumeration of the subscriptions registered on the local machine.
        /// </summary>
        /// <param name="SubscriptionEnum">The handle to the enumerator object that is returned from the EcOpenSubscriptionEnum function.</param>
        /// <param name="SubscriptionNameBufferSize">The size of the user-supplied buffer (in chars) to store the subscription name.</param>
        /// <param name="SubscriptionNameBuffer">The user-supplied buffer to store the subscription name.</param>
        /// <param name="SubscriptionNameBufferUsed">The size of the user-supplied buffer that is used by the function on successful return, or the size that is necessary to store the subscription name when the function fails with ERROR_INSUFFICIENT_BUFFER.</param>
        /// <returns>True if successful, False if the function failed. Use the GetLastError function to obtain the error code.</returns>
        [DllImport("wecapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool EcEnumNextSubscription(
            IntPtr SubscriptionEnum,
            Int32 SubscriptionNameBufferSize,
            IntPtr SubscriptionNameBuffer,
            ref Int32 SubscriptionNameBufferUsed
            );

        /// <summary>
        /// The EcGetObjectArrayProperty function retrieves property values from a handle to an array of event source properties. The array contains property values for the event sources of a subscription.
        /// </summary>
        /// <param name="ObjectArray">A handle to an array of properties for the event sources for a subscription. An array handle that is returned by the EcGetSubscriptionProperty method when the EcSubscriptionEventSources value is passed into the PropertyId parameter.</param>
        /// <param name="PropertyId">The property identifier for properties in the array. Specify a value from the EC_SUBSCRIPTION_PROPERTY_ID enumeration. Get the Address, Enabled, UserName, and Password properties in the array by specifying the EcSubscriptionEventSourceAddress, EcSubscriptionEventSourceEnabled, EcSubscriptionEventSourceUserName, or EcSubscriptionEventSourcePassword values.</param>
        /// <param name="ArrayIndex">The index of the array that specifies which event source to get the property from.</param>
        /// <param name="Flags">Reserved. Must be 0.</param>
        /// <param name="PropertyValueBufferSize">The size of the buffer that contains the value of the property. The size must be at least the size of an EC_VARIANT value.</param>
        /// <param name="PropertyValueBuffer">The user-supplied buffer to store property value into.</param>
        /// <param name="PropertyValueBufferUsed">The size of the user-supplied buffer that is used by the function on successful return, or the size that is necessary to store the property value when the</param>
        /// <returns></returns>
        [DllImport("wecapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool EcGetObjectArrayProperty(
            IntPtr ObjectArray,
            EC_SUBSCRIPTION_PROPERTY_ID PropertyId,
            Int32 ArrayIndex,
            Int32 Flags,
            Int32 PropertyValueBufferSize,
            IntPtr PropertyValueBuffer,
            ref Int32 PropertyValueBufferUsed
            );

        ///TODO:Implement Signature: EcGetObjectArraySize

        /// <summary>
        /// The EcGetSubscriptionProperty function retrieves a specific property value from a subscription object. The subscription object is specified by the handle passed into the Subscription parameter.
        /// </summary>
        /// <param name="Subscription">The handle to the subscription object.</param>
        /// <param name="PropertyId">An identifier that specifies which property of the subscription to get. Specify a value from the EC_SUBSCRIPTION_PROPERTY_ID enumeration. If you specify the EcSubscriptionEventSources value, then a handle to an array (EC_OBJECT_ARRAY_PROPERTY_HANDLE) will be returned. You can then use the EcGetObjectArrayProperty and EcSetObjectArrayProperty functions to get and set the Address, Enabled, UserName, and Password properties in the array</param>
        /// <param name="Flags">Reserved. Must be zero.</param>
        /// <param name="PropertyValueBufferSize">The size of the user-supplied buffer to store the property value into.</param>
        /// <param name="PropertyValueBuffer">The user-supplied buffer to store property value into.</param>
        /// <param name="PropertyValueBufferUsed">The size of the user-supplied buffer that is used by the function on successful return, or the size that is necessary to store the property value when function fails with ERROR_INSUFFICIENT_BUFFER.</param>
        /// <returns>True if the function was successful. False if the function failed. Use the GetLastError function to obtain the error code.</returns>
        [DllImport("wecapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool EcGetSubscriptionProperty(
            IntPtr Subscription,
            EC_SUBSCRIPTION_PROPERTY_ID PropertyId,
            Int32 Flags,
            Int32 PropertyValueBufferSize,
            IntPtr PropertyValueBuffer,
            ref Int32 PropertyValueBufferUsed
            );

        /// <summary>
        /// The EcGetSubscriptionRunTimeStatus function retrieves the run time status information for an event source of a subscription or the subscription itself. The subscription is specified by its name. If the event source is NULL, then the status for the overall subscription is retrieved.
        /// </summary>
        /// <param name="SubscriptionName"></param>
        /// <param name="StatusInfoId">An identifier that specifies which run time status information to get from the subscription. Specify a value from the EC_SUBSCRIPTION_RUNTIME_STATUS_INFO_ID enumeration. The EcSubscriptionRunTimeStatusEventSources value can be used to obtain the list of event sources associated with a subscription.</param>
        /// <param name="EventSourceName">The name of the event source to get the status from. Each subscription can have multiple event sources.</param>
        /// <param name="Flags">Reserved. Must be zero.</param>
        /// <param name="PropertyValueBufferSize">The size of the user-supplied buffer that will hold the run time status information.</param>
        /// <param name="PropertyValueBuffer">The user-supplied buffer that will hold the run time status information. The buffer will hold the appropriate value depending on the EC_SUBSCRIPTION_RUNTIME_STATUS_INFO_ID value passed into the StatusInfoId parameter.</param>
        /// <param name="PropertyValueBufferUsed">The size of the user supplied buffer that is used by the function on successful return, or the size that is necessary to store the property value when function fails with ERROR_INSUFFICIENT_BUFFER.</param>
        /// <returns>True if successful, false if not. Use GetLastError to obtain the error code.</returns>
        [DllImport("wecapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool EcGetSubscriptionRunTimeStatus(
            string SubscriptionName,
            EC_SUBSCRIPTION_RUNTIME_STATUS_INFO_ID StatusInfoId,
            string EventSourceName,
            Int32 Flags,
            Int32 PropertyValueBufferSize,
            IntPtr PropertyValueBuffer,
            ref Int32 PropertyValueBufferUsed
            );

        ///TODO:Implement Signature: EcInsertObjectArrayElement  

        /// <summary>
        /// The EcOpenSubscriptionEnum function is creates a subscription enumerator to enumerate all registered subscriptions on the local machine.
        /// </summary>
        /// <param name="Flags">Reserved, must be zero.</param>
        /// <returns>If the function succeeds, it returns an handle (EC_HANDLE) to a new subscription enumerator object. Returns NULL otherwise, in which case use the GetLastError function to obtain the error code.</returns>
        [DllImport("wecapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr EcOpenSubscriptionEnum(
            Int32 Flags
            );

        /// <summary>
        /// The EcOpenSubscription function is used to open an existing subscription or create a new subscription according to the flag value specified.
        /// </summary>
        /// <param name="SubscriptionName">Specifies the name of the subscription. The value provided for this parameter should be unique within the computer's scope.</param>
        /// <param name="AccessMask">An access mask that specifies the desired access rights to the subscription. Use the EC_READ_ACCESS or EC_WRITE_ACCESS constants to specify the access rights. The function fails if the security descriptor of the subscription does not permit the requested access for the calling process.</param>
        /// <param name="Flags">A value specifying whether a new or existing subscription will be opened. Use the EC_CREATE_NEW, EC_OPEN_ALWAYS, or EC_OPEN_EXISTING constants.</param>
        /// <returns>If the function succeeds, it returns an handle (EC_HANDLE) to a new subscription object. Returns NULL otherwise, in which case use the GetLastError function to obtain the error code.</returns>
        [DllImport("wecapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr EcOpenSubscription(
             string SubscriptionName,
             Int32 AccessMask,
             Int32 Flags
            );

        /// <summary>
        /// The EcSaveSubscription function saves subscription configuration information. This function should be called whenever new values are added or updated to the subscription by the EcSetSubscriptionProperty method. If the subscription is enabled, the subscription will be activated when it is saved.
        /// </summary>
        /// <param name="Subscription">The handle to the subscription object.</param>
        /// <param name="Flags">Reserved. Must be 0.</param>
        /// <returns>True if successful</returns>
        [DllImport("wecapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool EcSaveSubscription(
            IntPtr Subscription,
            Int32 Flags
            );

        /// <summary>
        /// The EcSetObjectArrayProperty function sets a property value in an array of property values for the event sources of a subscription.
        /// </summary>
        /// <param name="ObjectArray">A handle to the array that contains the property value to set. The array contains property values for the event sources of a subscription. The array handle is returned by the EcGetSubscriptionProperty method when the EcSubscriptionEventSources value is passed into the Subscription parameter.</param>
        /// <param name="PropertyId">An identifier that specifies which property to set. Specify a value from the EC_SUBSCRIPTION_PROPERTY_ID enumeration. Set the Address, Enabled, UserName, and Password properties in the array by specifying the EcSubscriptionEventSourceAddress, EcSubscriptionEventSourceEnabled, EcSubscriptionEventSourceUserName, or EcSubscriptionEventSourcePassword values.</param>
        /// <param name="ArrayIndex">The index of the object in the array to set a property value on.</param>
        /// <param name="Flags">Reserved. Must be 0.</param>
        /// <param name="PropertyValue">The value of the property.</param>
        /// <returns>TRUE The function was successful. FALSE The function failed.Use the GetLastError function to obtain the error code.</returns>
        [DllImport("wecapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool EcSetObjectArrayProperty(
            IntPtr ObjectArray,
            EC_SUBSCRIPTION_PROPERTY_ID PropertyId,
            Int32 ArrayIndex,
            Int32 Flags,
            IntPtr PropertyValue
            );

        /// <summary>
        /// The EcSetSubscriptionProperty function sets new values or updates existing values of a subscription. New values set through this method will not be active unless they are saved by the EcSaveSubscription method.
        /// </summary>
        /// <param name="Subscription">The handle to the subscription object.</param>
        /// <param name="PropertyId">A value from the EC_SUBSCRIPTION_PROPERTY_ID enumeration that specifies which property of the subscription to set.</param>
        /// <param name="Flags">Reserved. Must be 0.</param>
        /// <param name="PropertyValue">The value of the property to set for the indicated subscription property.</param>
        /// <returns>TRUE The function was successful. FALSE The function failed.Use the GetLastError function to obtain the error code.</returns>
        [DllImport("wecapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool EcSetSubscriptionProperty(
            IntPtr Subscription,
            Int32 PropertyId,
            Int32 Flags,
            IntPtr PropertyValue
            );




        ///TODO:Implement Signature: EcRemoveObjectArrayElement 

        ///TODO:Implement Signature: EcRetrySubscription 



    } // class NativeMethods
} // namespace WecAdmin
