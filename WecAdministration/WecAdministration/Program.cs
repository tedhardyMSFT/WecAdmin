﻿// This is the wrapper program for unit testing purposes only.
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
            // commented to see if initial call to enumerate subscriptions is causing heap corruption.
            // 2016-11-26:THIS DOES SEEM TO BE THE CASE - HEAP CORRUPTION NOT OBSERVED WHEN COMMENTED OUT!
            // things to check: verify signatures/structure definitions.
            // does the wecsvc tolarate having multiple handles open to it from same process?
            Console.WriteLine("enumerating all subscriptions.");
            List<string> subs = WecAdmin.EventCollectorAdmin.EnumerateSubscriptions();
            Console.WriteLine("done!");

            Console.WriteLine("Local subscriptions");
            for (int j = 0; j < subs.Count; j++)
            {
                Console.WriteLine("\tSubscription:{0}", subs[j]);
            }
            Console.WriteLine("end subscriptions");

            // hardcoded to work with top swubscriptionName.
            string subName = subs[0];
            Console.WriteLine("Getting event filter");
            string currentEventFilter = WecAdmin.EventCollectorAdmin.GetSubscriptionFilter(subName);
            Console.WriteLine("Filter:{0}", currentEventFilter);

            string NewEventFilter = "<QueryList><Query Id='0' Path='System'><Select Path='System'>*[System[(EventID=7045)]]</Select></Query></QueryList>";
            Console.WriteLine("Updating filter.");
            bool filterUpdate = WecAdmin.EventCollectorAdmin.SetSubscriptionFilter(subName, NewEventFilter);
            Console.WriteLine("Update status:{0}", filterUpdate);

            currentEventFilter = WecAdmin.EventCollectorAdmin.GetSubscriptionFilter(subName);
            Console.WriteLine("New Filter:{0}", currentEventFilter);

            Console.WriteLine("Getting sources for sub:{0}", subName);
            List<string> subSources = WecAdmin.EventCollectorAdmin.ListSubscriptionRuntimeEventSources(subName);

            Console.WriteLine("[WecAdmin]:Getting heartbeat times for eventsources for subscription:{0}", subName);
            for(int i = 0; i < subSources.Count; i++)
            {
                DateTime lastHeartbeat = WecAdmin.EventCollectorAdmin.GetEventSourceLastHeartbeat(subName, subSources[i]);
                string sourceStatus = WecAdmin.EventCollectorAdmin.GetEventSourceStatus(subName, subSources[i]);
                Console.WriteLine("\tSource:{0}\tHeartbeat:{1}\tStatus:{2}", subSources[i], lastHeartbeat.ToString("o"), sourceStatus);
                
            }
            Console.WriteLine("Done getting hearbeat times");
            Console.WriteLine("Hit Enter to exit.");
            //Console.ReadLine();
        } // static void Main(string[] args)
    } // class Program
}
