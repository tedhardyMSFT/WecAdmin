// This is the wrapper program for unit testing purposes only.
// see WecAdmin project for .NET to Win32 interface

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace WecAdministration
{
    // TODO:TOTAL:Feature Backlog:
    // DONE: Enumerate subscriptions (unfiltered)
    // DONE: enumerate subscription event sources (unfiltered)
    // DONE: enumearte subscription event sources status.
    // TODO: (optional flag for active and/or type: push/pull)
    // TODO: Get subscription parameter method (dest channel, heartbeat interval, account)
    //      One function that uses an enum for the property type.
    // TODO: Get subscription event sources (with flag for type: Active, inactive, all)
    // TODO: Create subscription function
    // TODO: Modify subscription property function (using enum for property type)
    // TODO: Delete sub function
    // TODO: Retry Sub function

    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("enumerating all subscriptions.");
            List<string> subs = WecAdmin.EventCollectorAdmin.EnumerateSubscriptions();
            Console.WriteLine("done!");
            string subName = string.Empty;
            Console.WriteLine("Local subscriptions");
            for (int j = 0; j < subs.Count; j++)
            {
                Console.WriteLine("\tSubscription:{0}", subs[j]);
            }
            Console.WriteLine("end subscriptions");

            if(subs.Count == 0 )
            {
                Console.WriteLine("no subscriptions. Exiting");
                Environment.Exit(1);
            }
            // hardcoded to work with top swubscriptionName.
            subName = subs[0];
            Console.WriteLine("Getting event filter");
            string currentEventFilter = WecAdmin.EventCollectorAdmin.GetSubscriptionFilter2(subName);
            Console.WriteLine("Filter:{0}", currentEventFilter);

            string NewEventFilter = "<QueryList><Query Id='0' Path='System'><Select Path='System'>*[System[(EventID=1945)]]</Select></Query></QueryList>";
            Console.WriteLine("Updating filter.");
            bool filterUpdate = WecAdmin.EventCollectorAdmin.SetSubscriptionFilter2(subName, NewEventFilter);
            Console.WriteLine("Update status:{0}", filterUpdate);
            //Console.WriteLine("Updating port.");
            bool portUpdate = WecAdmin.EventCollectorAdmin.SetSubscriptionPort(subName, 5985);
            //Console.WriteLine("Update status:{0}", portUpdate);
            currentEventFilter = WecAdmin.EventCollectorAdmin.GetSubscriptionFilter(subName);
            Console.WriteLine("New Filter:{0}", currentEventFilter);

            Console.WriteLine("Getting sources for sub:{0}", subName);
            List<string> subSources = WecAdmin.EventCollectorAdmin.ListSubscriptionRuntimeEventSources2(subName);

            Console.WriteLine("[WecAdmin]:Getting heartbeat times for eventsources for subscription:{0}", subName);
            for(int i = 0; i < subSources.Count; i++)
            {
                DateTime lastHeartbeat = WecAdmin.EventCollectorAdmin.GetEventSourceLastHeartbeat2(subName, subSources[i]);
                string sourceStatus = WecAdmin.EventCollectorAdmin.GetEventSourceStatus2(subName, subSources[i]);
                Console.WriteLine("\tSource:{0}\tHeartbeat:{1}\tStatus:{2}", subSources[i], lastHeartbeat.ToString("o"), sourceStatus);
                
            }

            Console.WriteLine("Done getting hearbeat times");

            WecAdmin.EventCollectorAdmin.SetSubscriptionContentFormat(subName, false);

            WecAdmin.EventCollectorAdmin.SetSubscriptionDestinationChannel(subName, "Application");

            Console.WriteLine("Hit Enter to exit.");
            Console.ReadLine();
        } // static void Main(string[] args)
    } // class Program
}
