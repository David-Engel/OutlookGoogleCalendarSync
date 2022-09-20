using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading;

namespace OutlookGoogleSync
{
    /// <summary>
    /// Description of GoogleCalendar.
    /// </summary>
    internal class GoogleCalendar
    {

        private static GoogleCalendar _instance;
        private static readonly string appName = Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyTitleAttribute>().Title;

        internal static GoogleCalendar Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new GoogleCalendar();
                return _instance;
            }
        }

        private readonly CalendarService service;

        internal GoogleCalendar()
        {
            UserCredential credential;

            using (var stream =
                new FileStream("gcs.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = @"token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    new string[] { CalendarService.ScopeConstants.CalendarReadonly, CalendarService.ScopeConstants.CalendarEvents },
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Google Calendar API service.
            service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = appName,
            });
        }

        internal List<GoogleCalendarListEntry> GetCalendars()
        {
            CalendarList request = service.CalendarList.List().Execute();

            if (request != null)
            {
                List<GoogleCalendarListEntry> result = new List<GoogleCalendarListEntry>();
                foreach (CalendarListEntry cle in request.Items)
                {
                    result.Add(new GoogleCalendarListEntry(cle));
                }
                return result;
            }
            return null;
        }

        internal List<Event> GetCalendarEntriesInRange(DateTime syncDateTime)
        {
            List<Event> result = new List<Event>();
            EventsResource.ListRequest lr = service.Events.List(Settings.Instance.SelectedGoogleCalendar.Id);

            lr.TimeMin = syncDateTime.AddDays(-Settings.Instance.DaysInThePast);
            lr.TimeMax = syncDateTime.AddDays(+Settings.Instance.DaysInTheFuture + 1);
            lr.SingleEvents = true;
            lr.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

            Events request = lr.Execute();

            // Fetch all pages of events
            if (request != null &&
                request.Items != null)
            {
                result.AddRange(request.Items);
                while (request.NextPageToken != null)
                {
                    lr.PageToken = request.NextPageToken;
                    request = lr.Execute();
                    result.AddRange(request.Items);
                }
            }

            return result;
        }

        internal void DeleteCalendarEntry(Event e)
        {
            service.Events.Delete(Settings.Instance.SelectedGoogleCalendar.Id, e.Id).Execute();
        }

        internal void AddEntry(Event e)
        {
            service.Events.Insert(e, Settings.Instance.SelectedGoogleCalendar.Id).Execute();
        }
    }
}
