using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
/*
Deutsch:
Die Funktion sucht in einem Outlook Kalender nach bestimmten Suchstrings und speichert das Ergebnis in eine Liste.
Es kann eine Ignore Liste eingegeben werden und oder mehrere Alias können für einen Sucherstringkategorie angegeben werden.
Auch wird ein prozentualer Vergleich der Trefferquote erstellt.

English:
The function searches for specific strings in an Outlook calendar and saves the result in a list.
An ignore list can be entered and or multiple aliases can be specified for a search string category.
A percentage comparison of the hit rate is also created.
*/



namespace OutlookDienste
{
    #region Helper class
    public class SearchItem
    {
        public SearchItem() { FindName = string.Empty; Count = 0; OnWeekendCount = 0; ResultList = new List<FindIn>(); StringList = new List<string>(); }
        public string FindName;
        public List<string> StringList;
        public int Count;
        public int OnWeekendCount;
        public List<FindIn> ResultList;
    }
    public class FindIn
    {
        public FindIn() { Subject = string.Empty; Start = DateTime.MinValue; End = DateTime.MinValue; AllDayEvent = false; Weekend = false; TotalDays = 0; }
        public string Subject;
        public DateTime Start;
        public DateTime End;
        public bool AllDayEvent;
        public bool Weekend;
        public int TotalDays;
    }
    #endregion
    /// <summary>
    /// The function searches for specific search strings in an Outlook calendar
    /// </summary>
    public class Find_In_Outlook
    {
        /// <summary>
        /// The function searches for specific search strings in an Outlook calendar
        /// </summary>
        /// <param name="calendarName">Name of the calendar in which to search</param>
        /// <param name="startDate">Start date</param>
        /// <param name="endDate">End date</param>
        /// <param name="FindItems">List of strings to be searched for. The result is also saved in the list.</param>
        /// <param name="Filter">A basic filter, only the dates are used, in which the filter string occurs.</param>
        /// <param name="ignoreList">A list of words not searched</param>
        /// <param name="OnlyAllDayEvents">Appointments that last a full day only</param>
        /// <returns></returns>
        public static string FindAppointmentsInTheCalendar(
            string calendarName,
            DateTime startDate,
            DateTime endDate,
            List<SearchItem> FindItems,
            string Filter = "",
            List<string> ignoreList = null,
            bool OnlyAllDayEvents = false
            )
        {
            int countFromMin = 10;
            string result = string.Empty;


            if (startDate > endDate || string.IsNullOrWhiteSpace(calendarName)|| FindItems == null || FindItems.Count ==0) 
            {                
                return "ERROR: values?";
            }
            
            Outlook.Application OutlookApplication = new Outlook.Application();            
            NameSpace mapiNamespace = OutlookApplication.GetNamespace("MAPI");
            string handleError = string.Empty;
            
            //Reset
            foreach (SearchItem item in FindItems)
            {
                item.Count = 0;
                item.ResultList.Clear();
            }

            Recipient recipient = mapiNamespace.CreateRecipient(calendarName);
            recipient.Resolve();

            //Was the calendar found?
            if (!recipient.Resolved) return "ERROR: calendar not found!";

            Microsoft.Office.Interop.Outlook.Folder calFolder;
            calFolder = mapiNamespace.GetSharedDefaultFolder(recipient, OlDefaultFolders.olFolderCalendar) as Outlook.Folder;

            //calFolder = OutlookApplication.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar) as Outlook.Folder;

            try
            {
                //Is the appointment on a weekend?
                int StartIntoTheWeekend = (int)System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.FirstDayOfWeek + 4;

                //Load all appointments in the period
                Outlook.Items rangeAppts = Find_In_Outlook.GetAppointmentsInRange(calFolder, startDate, endDate);
                
                
                if (rangeAppts != null)
                {
                    foreach (Outlook.AppointmentItem appt in rangeAppts)
                    {
                        
                        string Subject = appt.Subject.ToUpper();
                        
                        if (string.IsNullOrWhiteSpace(Filter)|| Subject.Contains(Filter.Trim().ToUpper()))
                        {
                            #region Ignore List
                            if (ignoreList != null && ignoreList.Count > 0)
                            {
                                foreach (string item in ignoreList)
                                {
                                    if (Subject.Contains(item.Trim().ToUpper())) continue;
                                }
                            } 
                            #endregion

                            if (OnlyAllDayEvents && !appt.AllDayEvent) continue;
                            
                            Debug.WriteLine("Subject: " + appt.Subject + " Start: " + appt.Start.ToString("g") + " end:" + appt.End.ToString("g"));

                            bool handle = false;
                            foreach (SearchItem item in FindItems)
                            {                                
                                string actual = string.Empty;
                                //find only one string
                                if (item.StringList == null || item.StringList.Count == 0)
                                {
                                    actual = item.FindName;
                                    handle = FindPart(StartIntoTheWeekend, appt, handle, Subject, item, actual);
                                }
                                else
                                //find from list
                                {
                                    foreach (string fitem in item.StringList)
                                    {
                                        actual = fitem;
                                        handle = FindPart(StartIntoTheWeekend, appt, handle, Subject, item, actual);
                                        //A value from the list was found in an appointment. Only one hit is counted
                                        if (handle) break;
                                    }
                                }
                                if (handle) break;
                            }

                            if (!handle && !string.IsNullOrWhiteSpace(Filter))
                            {
                                handleError += "? Filter: " + Filter + "  Subject:" + appt.Subject + Environment.NewLine;
                            }
                        }
                    }
                }



                
                //name max len
                int maxlen = FindItems.Max(s => s.FindName.Length);

                //sort
                FindItems = FindItems.OrderByDescending(x => x.Count).ToList();
                //FindItems = FindItems.OrderByDescending(x => x.String).ToList();

                //Min Max
                int maxCount = FindItems.Max(s => s.Count);
                int minCount = FindItems.Where(s => s.Count > countFromMin).Min(s => s.Count);
                int WeekendmaxCount = FindItems.Max(s => s.OnWeekendCount);
                int WeekendminCount = FindItems.Where(s => s.OnWeekendCount > countFromMin).Min(s => s.OnWeekendCount);


                foreach (SearchItem item in FindItems)
                {
                    int len = item.FindName.Length;
                    string space = string.Empty;
                    string resultPart = string.Empty;
                    //fill space
                    if (len < maxlen)
                    {
                        for (int i = 0; i < (maxlen - len); i++)
                        {
                            space += " ";
                        }
                    }
                    
                    double diff = 0; string diffSrg = string.Empty;                    
                    if (item.Count > 1 && item.Count < maxCount)
                    {
                        diff = Math.Round(100 - (item.Count / (maxCount * 0.01)), 2);
                        diffSrg = " diff. -" + diff.ToString() + "%";
                    }
                    else if (item.Count > 1)
                    {
                        diffSrg = " diff. 100%";
                    }

                    resultPart = item.FindName + ": " + space + item.Count.ToString() + diffSrg + "   Weekend:" + (item.OnWeekendCount).ToString();
                    len = resultPart.Length;
                    
                    //fill space 2
                    space = "";
                    int maxlen2 = 50;
                    if (len < maxlen2)
                    {
                        for (int i = 0; i < (maxlen2 - len); i++)
                        {
                            space += " ";
                        }
                    }

                    double Wdiff = 0; string WdiffSrg = string.Empty;
                    if (item.OnWeekendCount > 1 && item.OnWeekendCount < WeekendmaxCount)
                    {
                        Wdiff = Math.Round(100 - (item.OnWeekendCount / (WeekendmaxCount * 0.01)), 2);
                        WdiffSrg = " diff. -" + Wdiff.ToString() + "%";
                    }
                    else if (item.OnWeekendCount > 1)
                    {
                        WdiffSrg = " diff. 100%";
                    }

                    result += resultPart + space + WdiffSrg + Environment.NewLine;
                }


                if (!string.IsNullOrEmpty(result))
                {
                    if (!string.IsNullOrEmpty(handleError))
                    {
                        handleError = "ERROR: not found:" + Environment.NewLine + handleError;
                    }
                    result += Environment.NewLine + handleError;
                }
            }
            catch (System.Exception ex)
            {
                result += "EXEC ERROR:" + ex.Message;
            }
            OutlookApplication = null;
            return result;
        }

        private static bool FindPart(int StartIntoTheWeekend, AppointmentItem appt, bool handle, string Subject, SearchItem item, string actual)
        {
            if (Subject.Contains(actual.ToUpper()))
            {
                int totalDays = (int)(appt.End - appt.Start).TotalDays;
                bool onWeekend = false;

                if ((int)appt.Start.DayOfWeek > StartIntoTheWeekend || (int)appt.End.DayOfWeek > StartIntoTheWeekend || ((int)appt.Start.DayOfWeek > (int)appt.End.DayOfWeek))
                {
                    item.OnWeekendCount += totalDays;
                    onWeekend = true;
                }
                else
                {
                    item.Count += totalDays;
                }

                item.ResultList.Add(new FindIn { Subject = appt.Subject, Start = appt.Start, End = appt.End, AllDayEvent = appt.AllDayEvent, TotalDays = totalDays, Weekend = onWeekend });
                handle = true;
            }

            return handle;
        }

        /// <summary>
        /// Get recurring appointments in date range.
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <returns>Outlook.Items</returns>
        private  static Outlook.Items GetAppointmentsInRange(
            Outlook.Folder folder, DateTime startDate, DateTime endDate)
        {
            string filter = "[Start] >= '"
                + startDate.ToString("g")
                + "' AND [End] <= '"
                + endDate.ToString("g") + "'";
            
            Debug.WriteLine(filter);
            
            try
            {
                Outlook.Items calItems = folder.Items;
                calItems.IncludeRecurrences = true;
                calItems.Sort("[Start]", Type.Missing);
                Outlook.Items restrictItems = calItems.Restrict(filter);
                if (restrictItems.Count > 0)
                {
                    return restrictItems;
                }
                else
                {
                    return null;
                }
            }
            catch { return null; }
        }
    }
}
