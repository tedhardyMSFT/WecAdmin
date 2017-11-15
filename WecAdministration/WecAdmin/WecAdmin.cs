using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using WecAdmin;
using System.Security.Permissions;
using System.Runtime.Serialization;

//TODO:DEVNOTE - look at SafeHandle and SafeBuffer instead of IntPtr.

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

    public class EventCollectorAdmin
    {
        /// <summary>
        /// The data area passed to a system call is too small.
        /// </summary>
        private const int ERROR_INSUFFICIENT_BUFFER = 0x7a;

        /// <summary>
        /// [WinError.h] No more data is available.
        /// </summary>
        private const int ERROR_NO_MORE_ITEMS = 259;

        // constants from: https://msdn.microsoft.com/en-us/library/aa385773(v=vs.85).aspx
        /// <summary>
        /// Read access control permission that allows information to be read from the event collector.
        /// </summary>
        private const Int32 EC_READ_ACCESS = 1;
        /// <summary>
        /// Write access control permission that allows information to be written to the event collector.
        /// </summary>
        private const Int32 EC_WRITE_ACCESS = 2;
        /// <summary>
        /// Opens an existing subscription or creates the subscription if it does not exist. Used by the EcOpenSubscription method.
        /// </summary>
        private const Int32 EC_OPEN_ALWAYS = 0;
        /// <summary>
        /// A flag passed to the EcOpenSubscription function specifying that a new subscription should be created.
        /// </summary>
        private const Int32 EC_CREATE_NEW = 1;
        /// <summary>
        /// A flag passed to the EcOpenSubscription function specifying that an existing subscription should be opened.
        /// </summary>
        private const Int32 EC_OPEN_EXISTING = 2;

        /// <summary>
        /// Used to mask out the array bit from the Type property of an EC_VARIANT to extract the type of the variant value.
        /// </summary>
        private const Int32 EC_VARIANT_TYPE_MASK = 0x7f;

        /// <summary>
        /// When this bit is set in the Type property of an EC_VARIANT, the variant contains a pointer to an array of values, rather than the value itself.
        /// </summary>
        private const Int32 EC_VARIANT_TYPE_ARRAY = 0x80;

        /// <summary>
        /// Returns a list of the names of subscriptions registered on the local system.
        /// </summary>
        /// <returns>List of subscription names.</returns>
        public static List<string> EnumerateSubscriptions()
        {
            // output list of subscription names   
            List<string> SubscriptionList  = new List<string>();
            bool enumReturnVal = true;
            IntPtr ecEnumHandle = IntPtr.Zero;
            int lastWin32Error = -1;

            // Open Subscription enum handle
            // if zero, then return empty list.
            ecEnumHandle = NativeMethods.EcOpenSubscriptionEnum(0);
            if (IntPtr.Zero == ecEnumHandle)
            {
                lastWin32Error =  Marshal.GetLastWin32Error();
                // no handle returned for enumerating subscriptions.
                throw new EventCollectorApiException("Error opening subscription enumeration handle", lastWin32Error, "EcOpenSubscriptionEnum");
            }
            
            // keep unumerating until no more subscriptions.
            while (enumReturnVal)
            {
                Int32 bufferSize = 0;
                Int32 bufferUsed = 0;
                IntPtr outputBuffer = IntPtr.Zero;
                // this call will fail due to insufficient buffer
                enumReturnVal = NativeMethods.EcEnumNextSubscription(
                    ecEnumHandle,
                    bufferSize,
                    outputBuffer,
                    ref bufferUsed);

                // get status from function
                lastWin32Error = Marshal.GetLastWin32Error();
                // reached end of list
                if (lastWin32Error == ERROR_NO_MORE_ITEMS)
                {
                    // no more subscriptions to enumerate.
                    break;
                }
                // need to allcate buffer
                if (lastWin32Error == ERROR_INSUFFICIENT_BUFFER)
                {
                    // this increases the buffer size since the memory byte count 
                    // needs to accomodate for a UNICODE string.
                    // allocate unmanaged buffer and resubmit call.
                    IntPtr allocPtr = Marshal.AllocHGlobal(bufferUsed * sizeof(char));
                    bufferSize = bufferUsed;
                    enumReturnVal = NativeMethods.EcEnumNextSubscription(
                        ecEnumHandle,
                        bufferSize,
                        allocPtr,
                        ref bufferUsed);
                    // if successful, add to list
                    if (enumReturnVal)
                    {
                        string subscriptionName = Marshal.PtrToStringAuto(allocPtr);
                        SubscriptionList.Add(subscriptionName);
                    }
                    // free unmanaged memory allocation
                    Marshal.FreeHGlobal(allocPtr);
                } // if (lastWin32Error == ERROR_INSUFFICIENT_BUFFER)
            } // while (enumReturnVal)
            // close the handle 
            NativeMethods.EcClose(ecEnumHandle);
            return SubscriptionList;
        } // public static List<string> EnumerateSubscriptions()

        /// <summary>
        /// Retrieves the last heartbeat or status update time for the event source from the subscription.
        /// </summary>
        /// <param name="SubscriptionName">Name of the subscription to enumerate</param>
        /// <param name="EventSourceName">Name of the event source to retreieve heartbeat status</param>
        /// <returns>DateTime of the latest heartbeat for the source. If no source then 1600-01-01</returns>
        public static DateTime GetEventSourceLastHeartbeat(string SubscriptionName, string EventSourceName)
        {
            int bufferSize = 0;
            int bufferUsed = 0;
            IntPtr outputBuffer = IntPtr.Zero;
            DateTime lastHeartbeat = DateTime.FromFileTimeUtc(0);

            bool getProp = NativeMethods.EcGetSubscriptionRunTimeStatus(
                SubscriptionName,
                NativeMethods.EC_SUBSCRIPTION_RUNTIME_STATUS_INFO_ID.EcSubscriptionRunTimeStatusLastHeartbeatTime,
                EventSourceName, 
                0, // pass in zero - docs say pass in NULL (is reserved)
                bufferSize,
                outputBuffer,
                ref bufferUsed
                );
            int lastError = Marshal.GetLastWin32Error();

            // insufficient buffer, expected, so re-run with proper buffer size
            if (lastError == ERROR_INSUFFICIENT_BUFFER)
            {
                // now know that we need a buffer of correct size.
                // alloc the required memory in unmanaged space
                IntPtr allocPtr = IntPtr.Zero;
                // this will throw on Out of Memory condition.
                allocPtr = Marshal.AllocHGlobal(bufferUsed);

                // Marshals data from a managed object to an unmanaged block of memory.
                //Marshal.StructureToPtr(outputBuffer, allocPtr, false);
                bufferSize = bufferUsed;
                getProp = NativeMethods.EcGetSubscriptionRunTimeStatus(
                    SubscriptionName,
                    NativeMethods.EC_SUBSCRIPTION_RUNTIME_STATUS_INFO_ID.EcSubscriptionRunTimeStatusLastHeartbeatTime,
                    EventSourceName, // pass in null for all sources.
                    0,
                    bufferSize,
                    allocPtr,
                    ref bufferUsed
                    );

                if (getProp)
                {
                    WecAdmin.EC_VARIANT results = Marshal.PtrToStructure<WecAdmin.EC_VARIANT>(allocPtr);
                    if (results.Type == (int)NativeMethods.EC_VARIANT_TYPE.EcVarTypeDateTime)
                    {
                        lastHeartbeat = DateTime.FromFileTimeUtc(Marshal.ReadInt64(allocPtr));
                    }
                }
                Marshal.FreeHGlobal(allocPtr);
            } // if (lastError == ERROR_INSUFFICIENT_BUFFER)
            return lastHeartbeat;
        } // public static DateTime GetEventSourceLastHeartbeat(string SubscriptionName, string EventSourceName)

        public static DateTime GetEventSourceLastHeartbeat2(string SubscriptionName, string EventSourceName)
        {
            if (string.IsNullOrEmpty(EventSourceName))
            {
                throw new ArgumentNullException("EventSourceName requird for last heartbeat status.");
            }

            DateTime lastHeartbeat = DateTime.FromFileTimeUtc(0);
            IntPtr outputBuffer = IntPtr.Zero;
            int callStatus = ExecGetSubscriptionRuntimeStatus(
                SubscriptionName, 
                EventSourceName,  
                NativeMethods.EC_SUBSCRIPTION_RUNTIME_STATUS_INFO_ID.EcSubscriptionRunTimeStatusLastHeartbeatTime, 
                ref outputBuffer
                );

            if (callStatus == 0)
            {
                // api returned zero, time to inspect output buffer
                if (outputBuffer != IntPtr.Zero)
                {
                    WecAdmin.EC_VARIANT results = Marshal.PtrToStructure<WecAdmin.EC_VARIANT>(outputBuffer);
                    if (results.Type == (int)NativeMethods.EC_VARIANT_TYPE.EcVarTypeDateTime)
                    {
                        lastHeartbeat = DateTime.FromFileTimeUtc(Marshal.ReadInt64(outputBuffer));
                    }
                    Marshal.FreeHGlobal(outputBuffer);
                }
            } // if (callStatus != 0)
            return lastHeartbeat;
        } // public static DateTime GetEventSourceLastHeartbeat2(string SubscriptionName, string EventSourceName)

        /// <summary>
        /// Retrieves the runtime status of the event source (computer FQDN) for the subscription.
        /// </summary>
        /// <param name="SubscriptionName"></param>
        /// <param name="EventSourceName"></param>
        /// <returns></returns>
        public static string GetEventSourceStatus(string SubscriptionName, string EventSourceName)
        {
            int bufferSize = 0;
            int bufferUsed = 0;
            IntPtr outputBuffer = IntPtr.Zero;
            string eventSourceStatus = string.Empty;

            bool getProp = NativeMethods.EcGetSubscriptionRunTimeStatus(
                SubscriptionName,
                NativeMethods.EC_SUBSCRIPTION_RUNTIME_STATUS_INFO_ID.EcSubscriptionRunTimeStatusActive,
                EventSourceName, // specific event source name to get status.
                0, // pass in zero - docs say pass in NULL (is reserved)
                bufferSize,
                outputBuffer,
                ref bufferUsed
                );
            int lastError = Marshal.GetLastWin32Error();

            // insufficient buffer, expected, so re-run with proper buffer size
            if (lastError == ERROR_INSUFFICIENT_BUFFER)
            {
                // now know that we need a buffer of correct size.
                // alloc the required memory in unmanaged space
                IntPtr allocPtr = IntPtr.Zero;
                //TODO:ErrorHandling
                allocPtr = Marshal.AllocHGlobal(bufferUsed);

                // Marshals data from a managed object to an unmanaged block of memory.
                //Marshal.StructureToPtr(outputBuffer, allocPtr, false);
                bufferSize = bufferUsed;
                getProp = NativeMethods.EcGetSubscriptionRunTimeStatus(
                    SubscriptionName,
                    NativeMethods.EC_SUBSCRIPTION_RUNTIME_STATUS_INFO_ID.EcSubscriptionRunTimeStatusActive,
                    EventSourceName, // pass in null for all sources.
                    0,
                    bufferSize,
                    allocPtr,
                    ref bufferUsed
                    );

                if (getProp)
                {
                    WecAdmin.EC_VARIANT results = Marshal.PtrToStructure<WecAdmin.EC_VARIANT>(allocPtr);
                    //// heartbeat (if present) is in FileTimeUTC format.
                    //Console.WriteLine("variant type: {0}", results.Type);
                    if (results.Type == (int)NativeMethods.EC_VARIANT_TYPE.EcVarTypeUInt32)
                    {
                        Int32 statusValue = (Marshal.ReadInt32(allocPtr));
                        switch (statusValue)
                        {
                            case 1:
                                eventSourceStatus = "Disabled";
                                break;
                            case 2:
                                eventSourceStatus = "Active";
                                break;

                            case 3:
                                eventSourceStatus = "Inactive";
                                break;
                            case 4:
                                eventSourceStatus = "Trying";
                                break;
                            default:
                                eventSourceStatus = "InvalidEventSourceStatus";
                                break;
                        }
                    }

                    //Console.WriteLine("\tSource Name: {0}\t Last Heartbeat:{1}", EventSourceName, lastHeartbeat);
                }
                Marshal.FreeHGlobal(allocPtr);
            } // if (lastError == ERROR_INSUFFICIENT_BUFFER)

            return eventSourceStatus;
        } // public static DateTime GetEventSourceLastHeartbeat(string SubscriptionName, string EventSourceName)

        /// <summary>
        /// Retrieves the runtime status of the event source (computer FQDN) for the subscription.
        /// </summary>
        /// <param name="SubscriptionName"></param>
        /// <param name="EventSourceName"></param>
        /// <returns></returns>
        public static string GetEventSourceStatus2(string SubscriptionName, string EventSourceName)
        {
            IntPtr outputBuffer = IntPtr.Zero;
            string eventSourceStatus = string.Empty;

            int returnCode = ExecGetSubscriptionRuntimeStatus(
                SubscriptionName, 
                EventSourceName, 
                NativeMethods.EC_SUBSCRIPTION_RUNTIME_STATUS_INFO_ID.EcSubscriptionRunTimeStatusActive, 
                ref outputBuffer
                );

            // insufficient buffer, expected, so re-run with proper buffer size
            if (returnCode == 0)
            {
                    WecAdmin.EC_VARIANT results = Marshal.PtrToStructure<WecAdmin.EC_VARIANT>(outputBuffer);
                    //// heartbeat (if present) is in FileTimeUTC format.
                    //Console.WriteLine("variant type: {0}", results.Type);
                    if (results.Type == (int)NativeMethods.EC_VARIANT_TYPE.EcVarTypeUInt32)
                    {
                        Int32 statusValue = (Marshal.ReadInt32(outputBuffer));
                        switch (statusValue)
                        {
                            case 1:
                                eventSourceStatus = "Disabled";
                                break;
                            case 2:
                                eventSourceStatus = "Active";
                                break;

                            case 3:
                                eventSourceStatus = "Inactive";
                                break;
                            case 4:
                                eventSourceStatus = "Trying";
                                break;
                            default:
                                eventSourceStatus = "InvalidEventSourceStatus";
                                break;
                        }
                    }
                Marshal.FreeHGlobal(outputBuffer);
            } // if (returnCode == 0)

            return eventSourceStatus;
        } // public static DateTime GetEventSourceLastHeartbeat(string SubscriptionName, string EventSourceName)


        /// <summary>
        /// Returns all event sources for the subscription name.
        /// </summary>
        /// <param name="subscriptionName"></param>
        /// <returns></returns>
        public static  List<string> ListSubscriptionRuntimeEventSources(string subscriptionName)
        {
            List<string> eventSources = new List<string>();
            //EC_VARIANT outputBuffer = new EC_VARIANT();
            IntPtr outputBuffer = IntPtr.Zero;
            Int32 bufferSize = 0;
            Int32 bufferUsed = 0;
            Console.WriteLine("initial call to retrieve sources");

            bool getProp = NativeMethods.EcGetSubscriptionRunTimeStatus(
                subscriptionName,
                NativeMethods.EC_SUBSCRIPTION_RUNTIME_STATUS_INFO_ID.EcSubscriptionRunTimeStatusEventSources,
                null,
                0, // pass in zero - docs say pass in NULL (is reserved)
                bufferSize,
                outputBuffer,
                ref bufferUsed
                );
            int lastError = Marshal.GetLastWin32Error();

            // insufficient buffer, expected, so re-run with proper buffer size
            if (lastError == ERROR_INSUFFICIENT_BUFFER)
            {
                Console.WriteLine("Retrying with buffer size:{0}", bufferUsed);
                // now know that we need a buffer of correct size.
                // alloc the required memory in unmanaged space
                IntPtr allocPtr = Marshal.AllocHGlobal(bufferUsed);
                // Marshals data from a managed object to an unmanaged block of memory.
                //Marshal.StructureToPtr(outputBuffer, allocPtr, false);
                bufferSize = bufferUsed;
                getProp = NativeMethods.EcGetSubscriptionRunTimeStatus(
                    subscriptionName,
                    NativeMethods.EC_SUBSCRIPTION_RUNTIME_STATUS_INFO_ID.EcSubscriptionRunTimeStatusEventSources,
                    null,
                    0,
                    bufferSize,
                    allocPtr,
                    ref bufferUsed
                    );

                if (getProp)
                {
                    EC_VARIANT results = Marshal.PtrToStructure<EC_VARIANT>(allocPtr);
                    int length = (int)results.Count;

                    for (int i = 0; i < length; i++)
                    {
                        // [pointer arithmetic] to move along the array, pointer by pointer.
                        IntPtr newPtr = IntPtr.Add(results.StringArr, (IntPtr.Size * i)); //results.StringArr + (IntPtr.Size * i);
                        // for each pointer, read the string it points to.
                        string sourceName = Marshal.PtrToStringUni(Marshal.ReadIntPtr(newPtr));
                        eventSources.Add(sourceName);
                    } // for (int i = 0; i < length; i++)
                } // if (getProp)
                // free unmanaged memory allocation
                Marshal.FreeHGlobal(allocPtr);
            } // if (lastError == ERROR_INSUFFICIENT_BUFFER)

            return eventSources;
        } // private static  List<string> ListSubscriptionRuntimeEventSources(string SubscriptionName)


        /// <summary>
        /// Returns all event sources for the subscription name.
        /// </summary>
        /// <param name="subscriptionName"></param>
        /// <returns></returns>
        public static List<string> ListSubscriptionRuntimeEventSources2(string subscriptionName)
        {
            List<string> eventSources = new List<string>();
            //EC_VARIANT outputBuffer = new EC_VARIANT();
            IntPtr outputBuffer = IntPtr.Zero;
            Console.WriteLine("initial call to retrieve sources");
            Int32 returnCode = ExecGetSubscriptionRuntimeStatus(subscriptionName, null, NativeMethods.EC_SUBSCRIPTION_RUNTIME_STATUS_INFO_ID.EcSubscriptionRunTimeStatusEventSources, ref outputBuffer);

            // insufficient buffer, expected, so re-run with proper buffer size
            if (returnCode == 0)
            {

                    EC_VARIANT results = Marshal.PtrToStructure<EC_VARIANT>(outputBuffer);
                    int length = (int)results.Count;

                    for (int i = 0; i < length; i++)
                    {
                        // [pointer arithmetic] to move along the array, pointer by pointer.
                        IntPtr newPtr = IntPtr.Add(results.StringArr, (IntPtr.Size * i)); //results.StringArr + (IntPtr.Size * i);
                        // for each pointer, read the string it points to.
                        string sourceName = Marshal.PtrToStringUni(Marshal.ReadIntPtr(newPtr));
                        eventSources.Add(sourceName);
                    } // for (int i = 0; i < length; i++)
                // free unmanaged memory allocation
                Marshal.FreeHGlobal(outputBuffer);
            } // if (lastError == ERROR_INSUFFICIENT_BUFFER)

            return eventSources;
        } // private static  List<string> ListSubscriptionRuntimeEventSources(string SubscriptionName)


        /// <summary>
        /// Retrieves the event query filter for the named subscription.
        /// </summary>
        /// <param name="subscriptionName">Name of the subscription</param>
        /// <returns>Event Query filter</returns>
        public static string GetSubscriptionFilter(string subscriptionName)
        {
            string eventFilter = string.Empty;

            // open subscription
            // read subscription event filter property
            // close subscription
            IntPtr subscriptionHandle = IntPtr.Zero;

            subscriptionHandle = OpenSubscription(subscriptionName, (int)EC_READ_ACCESS, (int)EC_OPEN_EXISTING);

            if (subscriptionHandle == IntPtr.Zero)
            {
                // throw here - subscription doesn't exist.
                Console.WriteLine("Subscription does not exist: {0}", subscriptionName);
                return string.Empty;
            }

            IntPtr outputBuffer = IntPtr.Zero;
            Int32 bufferSize = 0;
            Int32 bufferUsed = 0;
            // this will always fail
            bool getSubProperty = NativeMethods.EcGetSubscriptionProperty(
                subscriptionHandle,
                NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionQuery,
                0,
                bufferSize,
                outputBuffer,
                ref bufferUsed);

            int lastError = Marshal.GetLastWin32Error();

            // insufficient buffer, expected, so re-run with proper buffer size
            if (lastError == ERROR_INSUFFICIENT_BUFFER)
            {
                bufferSize = bufferUsed;
                Console.WriteLine("Retrying with buffer size:{0}", bufferUsed);
                // now know that we need a buffer of correct size.
                // alloc the required memory in unmanaged space
                IntPtr allocPtr = Marshal.AllocHGlobal(bufferUsed);
                // Marshals data from a managed object to an unmanaged block of memory.
                //Marshal.StructureToPtr(outputBuffer, allocPtr, false);
                bufferSize = bufferUsed;
                getSubProperty = NativeMethods.EcGetSubscriptionProperty(
                    subscriptionHandle,
                    NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionQuery,
                    0,
                    bufferSize,
                    allocPtr,
                    ref bufferUsed);

                if (getSubProperty)
                {
                    // convert into structure
                    WecAdmin.EC_VARIANT results = Marshal.PtrToStructure<WecAdmin.EC_VARIANT>(allocPtr);
                    // event Filter is a string type, read that value.
                    if (results.Type == (int)NativeMethods.EC_VARIANT_TYPE.EcVarTypeString)
                    {
                        eventFilter = Marshal.PtrToStringAuto(results.StringValue);
                    }
                } // if (getProp)
                // free unmanaged memory allocation
                Marshal.FreeHGlobal(allocPtr);

                // close the handle to the subscription.
                NativeMethods.EcClose(subscriptionHandle);
            } // if (lastError == ERROR_INSUFFICIENT_BUFFER)
            return eventFilter;
        } // public static string GetSubscriptionFilter(string subscriptionName)

        /// <summary>
        /// Retrieves the event query filter for the named subscription.
        /// </summary>
        /// <param name="subscriptionName">Name of the subscription</param>
        /// <returns>Event Query filter</returns>
        public static string GetSubscriptionFilter2(string subscriptionName)
        {
            string eventFilter = string.Empty;

            // open subscription
            // read subscription event filter property
            // close subscription
            IntPtr subscriptionHandle = IntPtr.Zero;
            IntPtr outputBuffer = IntPtr.Zero;

            subscriptionHandle = OpenSubscription(subscriptionName, (int)EC_READ_ACCESS, (int)EC_OPEN_EXISTING);

            int returnCode = ExecGetSubscriptionProperty(subscriptionHandle, NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionQuery, ref outputBuffer);

            // convert into structure
            WecAdmin.EC_VARIANT results = Marshal.PtrToStructure<WecAdmin.EC_VARIANT>(outputBuffer);
            // event Filter is a string type, read that value.
            if (results.Type == (int)NativeMethods.EC_VARIANT_TYPE.EcVarTypeString)
            {
                eventFilter = Marshal.PtrToStringAuto(results.StringValue);
            }

            // free unmanaged memory allocation
            Marshal.FreeHGlobal(outputBuffer);

            // close the handle to the subscription.
            NativeMethods.EcClose(subscriptionHandle);
            return eventFilter;
        } // public static string GetSubscriptionFilter(string subscriptionName)


        /// <summary>
        /// Updates the event query filter for the supplied subscription.
        /// </summary>
        /// <param name="SubscriptionName">Name of the subscription to update</param>
        /// <param name="EventFilter">new event query string</param>
        /// <returns>True if successful</returns>
        public static bool SetSubscriptionFilter(string SubscriptionName, string EventFilter)
        {
            string errorMessage = string.Empty;
            bool returnVal = false;
            // open handle to subscription with flags
            IntPtr subHandle = OpenSubscription(SubscriptionName, (int)(EC_READ_ACCESS | EC_WRITE_ACCESS), (int)EC_OPEN_EXISTING);
            // marshal string to unmanaged memory
            IntPtr filterPtr = Marshal.StringToHGlobalAuto(EventFilter);
            // allocate un-managed memory for the string and get the pointer.
            EC_VARIANT queryUpdate = new EC_VARIANT() { 
                // set the type in EC_VARIANT
                Type = (uint)NativeMethods.EC_VARIANT_TYPE.EcVarTypeString,
                StringValue = filterPtr
            };
            // get struct size and allocate un-managed memory
            int ecVariantSize = Marshal.SizeOf(queryUpdate);
            IntPtr ecVariantPtr = Marshal.AllocHGlobal(ecVariantSize);
            // marshal the pointer into un-managed memory
            Marshal.StructureToPtr<EC_VARIANT>(queryUpdate, ecVariantPtr, true);

            // make the Win32 call
            returnVal = NativeMethods.EcSetSubscriptionProperty(
                subHandle,
                (int)NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionQuery,
                0, // Flag is always null, per docs
                ecVariantPtr);
            int lastError = Marshal.GetLastWin32Error();
            errorMessage = new Win32Exception(lastError).Message;
            Console.WriteLine("update satus:{0} last error:{1} message:{2}", returnVal, lastError, errorMessage);
            Console.WriteLine("Saving subscription");
            // required for subscription changes to take effect.
            // May return non-zero status depending on subscription status.
            returnVal = NativeMethods.EcSaveSubscription(subHandle, 0);
            lastError = Marshal.GetLastWin32Error();
            errorMessage = new Win32Exception(lastError).Message;
            Console.WriteLine("update satus:{0} last error:{1} message:{2}", returnVal, lastError, errorMessage);
            // close the handle to the subscription.
            NativeMethods.EcClose(subHandle);

            //TODO: Verify if DestroyStructure is needed or not.
            //Marshal.DestroyStructure<EC_VARIANT>(ecVariantPtr);
            // free structure memory
            Marshal.FreeHGlobal(ecVariantPtr);
            // free event filter unmanaged memory
            Marshal.FreeHGlobal(filterPtr);
            return returnVal;
        } // public static bool SetSubscriptionFilter(string SubscriptionName, string EventFilter)


        /// <summary>
        /// Updates the event query filter for the supplied subscription.
        /// </summary>
        /// <param name="SubscriptionName">Name of the subscription to update</param>
        /// <param name="EventFilter">new event query string</param>
        /// <returns>True if successful</returns>
        public static bool SetSubscriptionFilter2(string SubscriptionName, string EventFilter)
        {
            string errorMessage = string.Empty;
            bool returnVal = false;
            int lastError = -1;
            // open handle to subscription with flags
            IntPtr subHandle = OpenSubscription(SubscriptionName, (int)(EC_READ_ACCESS | EC_WRITE_ACCESS), (int)EC_OPEN_EXISTING);

            lastError = ExecSetSubscriptionProperty(subHandle, NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionQuery, EventFilter);

            // required for subscription changes to take effect.
            // May return non-zero status depending on subscription status.
            returnVal = NativeMethods.EcSaveSubscription(subHandle, 0);
            lastError = Marshal.GetLastWin32Error();
            errorMessage = new Win32Exception(lastError).Message;
            Console.WriteLine("update satus:{0} last error:{1} message:{2}", returnVal, lastError, errorMessage);
            // close the handle to the subscription.
            NativeMethods.EcClose(subHandle);

            return returnVal;
        } // public static bool SetSubscriptionFilter(string SubscriptionName, string EventFilter)



        public static bool SetSubscriptionPort(string SubscriptionName, UInt32 PortNumber)
        {
            bool returnVal = false;
            string errorMessage = string.Empty;
            // open handle to subscription.
            IntPtr subHandle = OpenSubscription(SubscriptionName, (int)(EC_READ_ACCESS | EC_WRITE_ACCESS), (int)EC_OPEN_EXISTING);

            // allocate un-managed memory for the port and get the pointer.
            IntPtr portPtr = IntPtr.Zero;
            portPtr = Marshal.AllocHGlobal(sizeof(UInt32));
            Marshal.WriteInt32(portPtr, (int)PortNumber);

            EC_VARIANT portUpdate = new EC_VARIANT() {
                Type = (uint)NativeMethods.EC_VARIANT_TYPE.EcVarTypeUInt32,
                UInt32Val = portPtr
            };
            int ecVariantSize = Marshal.SizeOf(portUpdate);
            IntPtr ecVariantPtr = Marshal.AllocHGlobal(Marshal.SizeOf(portUpdate));
            Marshal.StructureToPtr(portUpdate, ecVariantPtr, true);

            returnVal = NativeMethods.EcSetSubscriptionProperty(
                subHandle,
                (Int32)NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionTransportPort,
                0,
                ecVariantPtr);

            Int32 lastError = Marshal.GetLastWin32Error();
            errorMessage = new Win32Exception(lastError).Message;
            Console.WriteLine("update satus:{0} last error:{1} message:{2}", returnVal, lastError, errorMessage);
            Console.WriteLine("Saving subscription");
            returnVal = NativeMethods.EcSaveSubscription(subHandle, 0);
            lastError = Marshal.GetLastWin32Error();
            errorMessage = new Win32Exception(lastError).Message;
            Console.WriteLine("update satus:{0} last error:{1} message:{2}", returnVal, lastError, errorMessage);

            // free unmanaged meory
            if (portPtr != IntPtr.Zero)
                Marshal.FreeHGlobal(portPtr);
            if (ecVariantPtr != IntPtr.Zero)
                Marshal.FreeHGlobal(ecVariantPtr);

            // close the handle to the subscription.
            NativeMethods.EcClose(subHandle);

            Console.WriteLine("update satus:{0} last error:{1}", returnVal, lastError);
            return returnVal;
        } // public static bool SetSubscriptionPort(string SubscriptionName, UInt32 PortNumber)

        public static bool SetSubscriptionContentFormat(string SubscriptionName, bool RenderedText)
        {
            bool returnVal = false;
            string errorMessage = string.Empty;
            // open handle to subscription.
            IntPtr subHandle = OpenSubscription(SubscriptionName, (int)(EC_READ_ACCESS | EC_WRITE_ACCESS), (int)EC_OPEN_EXISTING);
            NativeMethods.EC_SUBSCRIPTION_CONTENT_FORMAT contentFormat = NativeMethods.EC_SUBSCRIPTION_CONTENT_FORMAT.EcContentFormatEvents;
            if (RenderedText)
            {
                contentFormat = NativeMethods.EC_SUBSCRIPTION_CONTENT_FORMAT.EcContentFormatRenderedText;
            }
            IntPtr cfPtr = IntPtr.Zero;
            cfPtr = Marshal.AllocHGlobal(sizeof(UInt32));
            Marshal.WriteInt32(cfPtr, (int)contentFormat);

            EC_VARIANT subUpdate = new EC_VARIANT() {
                Type = (uint)NativeMethods.EC_VARIANT_TYPE.EcVarTypeUInt32,
                UInt32Val = cfPtr
            };
            int ecVariantSize = Marshal.SizeOf(subUpdate);
            IntPtr ecVariantPtr = Marshal.AllocHGlobal(Marshal.SizeOf(subUpdate));
            Marshal.StructureToPtr(subUpdate, ecVariantPtr, true);
            Console.WriteLine("Updating Subscription content format");
            returnVal = NativeMethods.EcSetSubscriptionProperty(
                subHandle,
                (Int32)NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionContentFormat,
                0,
                ecVariantPtr);

            Int32 lastError = Marshal.GetLastWin32Error();
            errorMessage = new Win32Exception(lastError).Message;
            Console.WriteLine("update satus:{0} last error:{1} message:{2}", returnVal, lastError, errorMessage);
            Console.WriteLine("Saving subscription content format");
            returnVal = NativeMethods.EcSaveSubscription(subHandle, 0);
            lastError = Marshal.GetLastWin32Error();
            errorMessage = new Win32Exception(lastError).Message;
            Console.WriteLine("update satus:{0} last error:{1} message:{2}", returnVal, lastError, errorMessage);

            // free structure memory
            if (cfPtr != IntPtr.Zero)
                Marshal.FreeHGlobal(cfPtr);
            if (ecVariantPtr != IntPtr.Zero)
                Marshal.FreeHGlobal(ecVariantPtr);

            // close the handle to the subscription.
            NativeMethods.EcClose(subHandle);
            return returnVal;
        } // public static bool SetSubscriptionContentFormat(string SubscriptionName, bool RenderedText)

        /// <summary>
        /// Updates the subscription to log to the supplied event channel name.
        /// </summary>
        /// <param name="SubscriptionName">Name of the subscription.</param>
        /// <param name="ChannelName"></param>
        /// <returns></returns>
        public static bool SetSubscriptionDestinationChannel(string SubscriptionName, string ChannelName)
        {
            string errorMessage = string.Empty;
            bool returnVal = false;
            // open handle to subscription with flags
            IntPtr subHandle = OpenSubscription(SubscriptionName, (int)(EC_READ_ACCESS | EC_WRITE_ACCESS), (int)EC_OPEN_EXISTING);
            // marshal string to unmanaged memory
            IntPtr filterPtr = Marshal.StringToHGlobalAuto(ChannelName);
            // allocate un-managed memory for the string and get the pointer.
            EC_VARIANT subUpdate = new EC_VARIANT()
            {
                // set the type in EC_VARIANT
                Type = (uint)NativeMethods.EC_VARIANT_TYPE.EcVarTypeString,
                StringValue = filterPtr
            };
            // get struct size and allocate un-managed memory
            int ecVariantSize = Marshal.SizeOf(subUpdate);
            IntPtr ecVariantPtr = Marshal.AllocHGlobal(ecVariantSize);
            // marshal the pointer into un-managed memory
            Marshal.StructureToPtr<EC_VARIANT>(subUpdate, ecVariantPtr, true);

            // make the Win32 call
            returnVal = NativeMethods.EcSetSubscriptionProperty(
                subHandle,
                (int)NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionLogFile,
                0, // Flag is always null, per docs
                ecVariantPtr);

            int lastError = Marshal.GetLastWin32Error();
            errorMessage = new Win32Exception(lastError).Message;
            Console.WriteLine("update satus:{0} last error:{1} message:{2}", returnVal, lastError, errorMessage);
            Console.WriteLine("Saving subscription");
            returnVal = NativeMethods.EcSaveSubscription(subHandle, 0);
            lastError = Marshal.GetLastWin32Error();
            errorMessage = new Win32Exception(lastError).Message;
            Console.WriteLine("update satus:{0} last error:{1} message:{2}", returnVal, lastError, errorMessage);



            // close the handle to the subscription.
            NativeMethods.EcClose(subHandle);
            // free structure memory
            Marshal.FreeHGlobal(ecVariantPtr);
            // free event filter unmanaged memory
            Marshal.FreeHGlobal(filterPtr);
            return returnVal;

        } // public static bool SetSubscriptionContentFormat(string SubscriptionName, bool RenderedText)

        /// <summary>
        /// Executes the win32 API EcGetSubscriptionProperty and returns the pointer to the buffer value.
        /// The calling code must release the unmanaged memory using Marshal.FreeHGlobal to avoid a memory leak.
        /// Throws a new EventCollectorApiException if not successful.
        /// </summary>
        /// <param name="subscriptionHandle"></param>
        /// <param name="SubscriptionProperty"></param>
        /// <param name="apiReturn"></param>
        /// <returns></returns>
        private static int ExecGetSubscriptionProperty(IntPtr subscriptionHandle, 
            NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID SubscriptionProperty, 
            ref IntPtr apiReturn)
        {
            if (subscriptionHandle == IntPtr.Zero)
            {
                // throw here - subscription doesn't exist.
                throw new ArgumentException("SubscriptionHandle must be initialized.");
            }

            Int32 bufferSize = 0;
            Int32 bufferUsed = 0;
            // this will always fail
            bool getSubProperty = NativeMethods.EcGetSubscriptionProperty(
                subscriptionHandle,
                SubscriptionProperty,
                0,
                bufferSize,
                apiReturn,
                ref bufferUsed);

            int lastError = Marshal.GetLastWin32Error();

            // insufficient buffer, expected, so re-run with proper buffer size
            if (lastError == ERROR_INSUFFICIENT_BUFFER)
            {
                bufferSize = bufferUsed;
                Console.WriteLine("Retrying with buffer size:{0}", bufferUsed);
                // now know that we need a buffer of correct size.
                // alloc the required memory in unmanaged space
                apiReturn = Marshal.AllocHGlobal(bufferUsed);
                // Marshals data from a managed object to an unmanaged block of memory.
                //Marshal.StructureToPtr(outputBuffer, allocPtr, false);
                bufferSize = bufferUsed;
                getSubProperty = NativeMethods.EcGetSubscriptionProperty(
                    subscriptionHandle,
                    SubscriptionProperty,
                    0,
                    bufferSize,
                    apiReturn,
                    ref bufferUsed);

                // update the latest return code
                lastError = Marshal.GetLastWin32Error();
            } // if (lastError == ERROR_INSUFFICIENT_BUFFER)

            if (!getSubProperty)
            {
                throw new EventCollectorApiException(
                    string.Format("Error executing EcGetSubscriptionProperty. Return code:{0}", lastError),
                    lastError,
                    "EcGetSubscriptionProperty");
            }
            return lastError;
        } // public static string GetSubscriptionFilter(string subscriptionName)


        /// <summary>
        /// Executes the win32 API EcGetSubscriptionRuntimeStatus and returns the pointer to the buffer value.
        /// The calling code must release the unmanaged memory using Marshal.FreeHGlobal to avoid a memory leak.
        /// Throws a new EventCollectorApiException if not successful.
        /// </summary>
        /// <param name="subscriptionName">(required) Name of the subscription to query</param>
        /// <param name="eventSourceName">(optional) if querying for a specific event source name. Pass in null for all event sources</param>
        /// <param name="StatusType">Status to retrieve</param>
        /// <param name="apiReturn">Pointer to unmanaged memmory allocated with return value. IntPtr.Zero indicates no return value.</param>
        /// <returns>The return code from the Win32 api</returns>
        private static int ExecGetSubscriptionRuntimeStatus(string subscriptionName,
            string eventSourceName,
            NativeMethods.EC_SUBSCRIPTION_RUNTIME_STATUS_INFO_ID StatusType,
            ref IntPtr apiReturn)
        {
            // subscription name is always required
            if (string.IsNullOrEmpty(subscriptionName))
            {
                throw new ArgumentNullException("subscriptionName parameter may not be null or whitespace");
            }

            int bufferSize = 0;
            int bufferUsed = 0;
            IntPtr outputBuffer = IntPtr.Zero;
            int lastError = -1;

            bool getProp = NativeMethods.EcGetSubscriptionRunTimeStatus(
               subscriptionName,
               StatusType,
               eventSourceName, // pass in null for all sources.
               0, // pass in zero - docs say pass in NULL (is reserved)
               bufferSize,
               outputBuffer,
               ref bufferUsed
               );
            lastError = Marshal.GetLastWin32Error();

            // insufficient buffer, This is expected and normal.
            // so re-run with proper buffer size
            if (lastError == ERROR_INSUFFICIENT_BUFFER)
            {
                // now know that we need a buffer of correct size.
                // alloc the required memory in unmanaged space
                // this will throw on Out of Memory condition.
                apiReturn = Marshal.AllocHGlobal(bufferUsed);

                // Marshals data from a managed object to an unmanaged block of memory.
                //Marshal.StructureToPtr(outputBuffer, allocPtr, false);
                bufferSize = bufferUsed;
                getProp = NativeMethods.EcGetSubscriptionRunTimeStatus(
                    subscriptionName,
                    StatusType,
                    eventSourceName, // pass in null for all sources.
                    0,
                    bufferSize,
                    apiReturn, // output now that buffer has been allocated
                    ref bufferUsed
                    );
                // update last error in case of other error - this should be zero (0)
                lastError = Marshal.GetLastWin32Error();
            } // if (lastError == ERROR_INSUFFICIENT_BUFFER)
            
            if (!getProp)
            {
                throw new EventCollectorApiException(
                    string.Format("Error executing EcGetSubscriptionRunTimeStatus. Return code:{0}", lastError),
                    lastError, "EcGetSubscriptionRunTimeStatus",
                    subscriptionName,
                    eventSourceName);
            }
            return lastError;
        } // private static int ExecGetSubscriptionRuntimeStatus (string subscriptionName, string eventSourceName, PInvokeMethods.EC_SUBSCRIPTION_RUNTIME_STATUS_INFO_ID StatusType)

        private static int ExecSetSubscriptionProperty(IntPtr SubscriptionHandle, NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID PropertyName, object value)
        {
            // TODO:IMPLEMENT
            /* use this page for reference https://msdn.microsoft.com/en-us/library/aa382725(v=vs.85).aspx
             * For each Property_ID value there is a single data-type associated with it.
             * use that to determine the cast of the object into the specific data type and then use that to craete a EC_VARIANT 
             * and then call EcSetSubscriptionProperty
             * 
             */
            
            EC_VARIANT updatedEcVariant = new EC_VARIANT();
            IntPtr ecVariantPtr = IntPtr.Zero;
            int ecVariantSize = 0;
            IntPtr updatedData = IntPtr.Zero;
            bool setApiReturnVal = false;
            int lastWin32ErrorCode = -1;


            // depending on the subscription property cast the value and set it in the structure
            switch (PropertyName)
            {
                // boolean types:
                case NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionEnabled:
                case NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionEventSourceEnabled:
                case NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionReadExistingEvents:
                    updatedEcVariant.Type = (int)NativeMethods.EC_VARIANT_TYPE.EcVarTypeBoolean;
                    // marshal here
                    throw new NotImplementedException("does not implement boolean types.");
                    break;
                // string array types
                case NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionEventSources:
                    // set type to be array
                    updatedEcVariant.Type = (int)NativeMethods.EC_VARIANT_TYPE.EcVarTypeString | NativeMethods.EC_VARIANT_TYPE_ARRAY;
                    // this is an array - need to marshal a pointer to an array of strings
                    throw new NotImplementedException("does not implement array types.");
                    break;
                // string types
                case NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionEventSourceAddress:
                case NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionEventSourceUserName:
                case NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionEventSourcePassword:
                case NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionDescription:
                case NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionURI:
                case NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionQuery:
                case NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionTransportName:
                case NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionLocale:
                case NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionLogFile:
                case NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionPublisherName:
                case NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionCommonUserName:
                case NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionCommonPassword:
                case NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionHostName:
                case NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionDialect:
                case NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionAllowedIssuerCAs:
                case NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionAllowedSubjects:
                case NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionDeniedSubjects:
                case NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID.EcSubscriptionAllowedSourceDomainComputers:
                    // marshal string into unmanaged memory
                    updatedData = Marshal.StringToHGlobalAuto((string)value);
                    // set fields on EC_VARIANT struct
                    updatedEcVariant.Type = (int)NativeMethods.EC_VARIANT_TYPE.EcVarTypeString;
                    updatedEcVariant.StringValue = updatedData;
                    // get size of struct  and allocate memory
                    ecVariantSize = Marshal.SizeOf(updatedEcVariant);
                    ecVariantPtr = Marshal.AllocHGlobal(ecVariantSize);
                    // marshal the structure into unmanaged memory
                    Marshal.StructureToPtr<EC_VARIANT>(updatedEcVariant, ecVariantPtr, true);
                    break;
            } // switch (PropertyName)


            // call the native function
            setApiReturnVal = NativeMethods.EcSetSubscriptionProperty(
                SubscriptionHandle,
                (int)PropertyName,
                0,
                ecVariantPtr
                );
            lastWin32ErrorCode = Marshal.GetLastWin32Error();

            //TODO: Verify if DestroyStructure is needed or not.
            //Marshal.DestroyStructure<EC_VARIANT>(ecVariantPtr);
            // clean up unmanaged memory if allocated
            if (ecVariantPtr != IntPtr.Zero)
            {
                Marshal.FreeHGlobal(ecVariantPtr);
            }
            if (updatedData != IntPtr.Zero)
            {
                Marshal.FreeHGlobal(updatedData);
            }

            if (!setApiReturnVal)
            {
                throw new EventCollectorApiException("Error calling EcSetSubscriptionProperty. Call failed with errorCode:"+ lastWin32ErrorCode, lastWin32ErrorCode, "EcSetSubscriptionProperty");
            }


            return lastWin32ErrorCode;
        } // private static int ExecSetSubscriptionProperty(IntPtr SubscriptionHandle, NativeMethods.EC_SUBSCRIPTION_PROPERTY_ID PropertyName, object value)

        /// <summary>
        /// Returns a handle to the subscription name supplied.
        /// This will be common across methods.
        /// Handle must be closed using the EcClose function.
        /// </summary>
        /// <param name="subscriptionName"></param>
        /// <returns></returns>
        private static IntPtr OpenSubscription(string subscriptionName,
            Int32 accessMask,
            Int32 flags)
        {
            IntPtr subHandle = IntPtr.Zero;
            subHandle = NativeMethods.EcOpenSubscription(
                subscriptionName,
                accessMask,
                flags);

            Int32 lastError = Marshal.GetLastWin32Error();

            if (lastError != 0 || subHandle == IntPtr.Zero)
            {
                throw new EventCollectorApiException(
                    string.Format("Error opening handle to Event Collector Subscription:{0}", subscriptionName),
                    lastError,
                    "EcOpenSubscription",
                    subscriptionName);
            }
            return subHandle;
        } // private static IntPtr openSubscription (string subscriptionName)
    } // public class EventCollectorAdmin
} // namespace WecAdmin
