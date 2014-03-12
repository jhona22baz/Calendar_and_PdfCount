using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Reflection;
using Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;



namespace Calendar
{
    class Program
    {
       
        private bool CreateCustomCalendar()
        {
            try
            {
                Outlook.Application olApp = (Outlook.Application)new Application();
                NameSpace mapiNS = olApp.GetNamespace("MAPI");

                string profile = "";
                mapiNS.Logon(profile, null, null, null);

                Outlook.AppointmentItem apt = (Outlook.AppointmentItem)olApp.CreateItem(OlItemType.olAppointmentItem);

                // set some properties
                apt.Subject = "MY Birthday";
                apt.Body = "This is an example of how to create a calendar event ";

                apt.Start = new DateTime(2014, 3, 8, 13, 30, 00);
                apt.End = new DateTime(2014, 3, 8, 14, 31, 00);
                apt.Importance = OlImportance.olImportanceHigh;
                apt.ReminderMinutesBeforeStart = 15;           // Number of minutes before the event for the remider
                apt.BusyStatus = OlBusyStatus.olTentative;    // Makes it appear bold in the calendar
                apt.AllDayEvent = false;
                apt.Location = "My house";
                apt.Save();
            }
            catch 
            {
                return false;
            }
            return true;
        }


        public bool viewAndModifyAppointment() 
        {
            try
            {
                // Create the Outlook application.
                Outlook.Application oApp = new Outlook.Application();

                // Get the NameSpace and Logon information.
                // Outlook.NameSpace oNS = (Outlook.NameSpace)oApp.GetNamespace("mapi");
                Outlook.NameSpace oNS = oApp.GetNamespace("mapi");

                //Log on by using a dialog box to choose the profile.
                oNS.Logon(Missing.Value, Missing.Value, true, true);

                //Alternate logon method that uses a specific profile.                
                //oNS.Logon("email", "pasword", false, true); 

                // Get the Calendar folder.
                Outlook.MAPIFolder oCalendar = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);

                // Get the Items (Appointments) collection from the Calendar folder.
                Outlook.Items oItems = oCalendar.Items;

                // Get the first item.                
                Outlook.AppointmentItem oAppt = (Outlook.AppointmentItem)oItems.Find("[Start] >= '15/03/2014 01:00 p. m.' AND [End] <= '15/03/2014 01:30 p. m.'");
                //"[Start] >= '15/03/2014 01:00 p. m.' AND [End] <= '15/03/2014 01:30 p. m.'"
                //"[Subject]='Revision Directiva Prueba 2'"

                /*
                DateTime start = DateTime.Parse("11/03/2014");
                DateTime end = start.AddDays(5);
                Console.WriteLine(start.ToString());
                start = oAppt.Start;
                end = oAppt.End;
                string filter = "[Start] >= '" + start.ToString("g") + "' AND [End] <= '" + end.ToString("g") + "'";
                Console.WriteLine(filter);
                */
                // Show some common properties.
                Console.WriteLine("Subject: " + oAppt.Subject);
                Console.WriteLine("Organizer: " + oAppt.Organizer);
                Console.WriteLine("Start: " + oAppt.Start.ToString());
                Console.WriteLine("End: " + oAppt.End.ToString());
                Console.WriteLine("Location: " + oAppt.Location);
                Console.WriteLine("Recurring: " + oAppt.IsRecurring);
                Console.WriteLine("Body: " + oAppt.Body);
                Console.WriteLine("Conversatiom ID: " + oAppt.ConversationID);
                Console.WriteLine("Duration: " + oAppt.Duration);
                
                //oAppt.Subject = "cumpleaños";
                //oAppt.Start = DateTime.Now.AddDays(5);
                //oAppt.End = DateTime.Now.AddDays(5);
                //oAppt.Body = "this is an example of how to change a calendar event.";

                /*
                    ["Subject"] 
                    ["Location"]
                    ["StartTime"]
                    ["EndTime"]
                    ["StartDate"] 
                    ["EndDate"] 
                    ["AllDayEvent"]
                    ["Body"] 
                */
                
                //oAppt.Send();//send item
                //oAppt.Save();//save item
                oAppt.Display(true);//Show the item to pause
                // Done. Log off.
                oNS.Logoff();

                // Clean up.
                oAppt = null;
                oItems = null;
                oCalendar = null;
                oNS = null;
                oApp = null;
            }
            //Simple error handling.
            catch (System.Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
                Console.Read();
            }
            return true;
        }
        

        public static int Main(string[] args)
        {
      
            //Default return value
            return 0;
        }
    }
}
