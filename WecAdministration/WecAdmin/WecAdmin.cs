using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using WecAdmin;


namespace WecAdmin
{
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
        /// <returns></returns>
        public static List<string> EnumerateSubscriptions()
        {
            List<string> SubscriptionList  = new List<string>();
            bool enumReturnVal = true;
            IntPtr ecEnumHandle = IntPtr.Zero;

            // Open Subscription enum handle
            // if zero, then return empty list.
            ecEnumHandle = PInvokeMethods.EcOpenSubscriptionEnum(0);
            if (IntPtr.Zero == ecEnumHandle)
            {
                // nothing to enumerate, return empty list.
                return SubscriptionList;
            }
            
            while (enumReturnVal)
            {
                Int32 bufferSize = 0;
                IntPtr buffer = IntPtr.Zero;
                // this call will fail due to insufficient buffer
                enumReturnVal = PInvokeMethods.EcEnumNextSubscription(
                    ecEnumHandle,
                    bufferSize,
                    buffer,
                    ref bufferSize);

                // get status from function
                Int32 statusMessage = Marshal.GetLastWin32Error();
                // reached end of list
                if (statusMessage == ERROR_NO_MORE_ITEMS)
                {
                    // end of list, return accumulated list so far
                    return SubscriptionList;
                }
                // need to allcate buffer
                if (statusMessage == ERROR_INSUFFICIENT_BUFFER)
                {
                    // allocate unmanaged buffer and resubmit call.
                    IntPtr allocPtr = Marshal.AllocHGlobal(bufferSize);
                    enumReturnVal = PInvokeMethods.EcEnumNextSubscription(
                        ecEnumHandle,
                        bufferSize,
                        allocPtr,
                        ref bufferSize);

                    if (enumReturnVal)
                    {
                        string subscriptionName = Marshal.PtrToStringAuto(allocPtr);
                        SubscriptionList.Add(subscriptionName);
                    }
                        
                    // free unmanaged memory allocation
                    Marshal.FreeHGlobal(allocPtr);
                }
            } // while (enumReturnVal)

            // close the handle
            PInvokeMethods.EcClose(ecEnumHandle);
            // Using handle call enumSubscription
            //  two cases:
            //      1) insufficient buffer - so call enumSubscription again with correct buffer size to get next subscription name
            //      2) no more items - end of list.
            return SubscriptionList;
        } // public static List<string> EnumerateSubscriptions()

        public static DateTime GetEventSourceLastHeartbeat(string SubscriptionName, string EventSourceName)
        {
            int bufferSize = 0;
            int bufferUsed = 0;
            IntPtr outputBuffer = IntPtr.Zero;
            DateTime lastHeartbeat = DateTime.FromFileTimeUtc(0);

            bool getProp = PInvokeMethods.EcGetSubscriptionRunTimeStatus(
                SubscriptionName,
                PInvokeMethods.EC_SUBSCRIPTION_RUNTIME_STATUS_INFO_ID.EcSubscriptionRunTimeStatusLastHeartbeatTime,
                EventSourceName, // pass in null for all sources.
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
                getProp = PInvokeMethods.EcGetSubscriptionRunTimeStatus(
                    SubscriptionName,
                    PInvokeMethods.EC_SUBSCRIPTION_RUNTIME_STATUS_INFO_ID.EcSubscriptionRunTimeStatusLastHeartbeatTime,
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
                    Console.WriteLine("variant type: {0}", results.Type);
                    if (results.Type == (int)PInvokeMethods.EC_VARIANT_TYPE.EcVarTypeDateTime)
                    {
                        lastHeartbeat = DateTime.FromFileTimeUtc(Marshal.ReadInt64(allocPtr));
                    }

                    Console.WriteLine("\tSource Name: {0}\t Last Heartbeat:{1}", EventSourceName, lastHeartbeat);
                }
                Marshal.FreeHGlobal(allocPtr);
            } // if (lastError == ERROR_INSUFFICIENT_BUFFER)

            return lastHeartbeat;
        } // public static DateTime GetEventSourceLastHeartbeat(string SubscriptionName, string EventSourceName)


        public static  List<string> ListSubscriptionRuntimeEventSources(string subscriptionName)
        {
            List<string> eventSources = new List<string>();
            //EC_VARIANT outputBuffer = new EC_VARIANT();
            IntPtr outputBuffer = IntPtr.Zero;
            Int32 bufferSize = 0;
            Int32 bufferUsed = 0;
            Console.WriteLine("initial call to retrieve sources");
            bool getProp = PInvokeMethods.EcGetSubscriptionRunTimeStatus(
                subscriptionName,
                PInvokeMethods.EC_SUBSCRIPTION_RUNTIME_STATUS_INFO_ID.EcSubscriptionRunTimeStatusEventSources,
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
                getProp = PInvokeMethods.EcGetSubscriptionRunTimeStatus(
                    subscriptionName,
                    PInvokeMethods.EC_SUBSCRIPTION_RUNTIME_STATUS_INFO_ID.EcSubscriptionRunTimeStatusEventSources,
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
                //Marshal.FreeHGlobal(allocPtr);
            } // if (lastError == ERROR_INSUFFICIENT_BUFFER)

            return eventSources;
        } // private static  List<string> ListSubscriptionRuntimeEventSources(string SubscriptionName)

        /// <summary>
        /// Returns a handle to the subscription name supplied.
        /// This will be common across methods.
        /// </summary>
        /// <param name="subscriptionName"></param>
        /// <returns></returns>
        private static IntPtr openSubscription (string subscriptionName)
        {
            IntPtr subHandle = IntPtr.Zero;

            subHandle = PInvokeMethods.EcOpenSubscription(
                subscriptionName,
                EC_READ_ACCESS,
                EC_OPEN_EXISTING
                );

            return subHandle;
        } // private static IntPtr openSubscription (string subscriptionName)
    } // public class EventCollectorAdmin
} // namespace WecAdmin
