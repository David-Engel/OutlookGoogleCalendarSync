using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Diagnostics;
using DotNetOpenAuth.OAuth2;
using Google.Apis.Authentication.OAuth2;
using Google.Apis.Authentication.OAuth2.DotNetOpenAuth;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Util;

namespace OutlookGoogleSync
{
    /// <summary>
    /// Description of GoogleCalendar.
    /// </summary>
    public class GoogleCalendar
    {

        private static GoogleCalendar _instance;

        public static GoogleCalendar Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new GoogleCalendar();
                return _instance;
            }
        }

        CalendarService service;

        public GoogleCalendar()
        {
            var provider = new NativeApplicationClient(GoogleAuthenticationServer.Description);
            provider.ClientIdentifier = "1048601141806-9ek0gtv3bjonb7cifqaedf8gh3e8sive.apps.googleusercontent.com";
            provider.ClientSecret = "F6uRAziIV-EdFsb6oRjTAY45";
            service = new CalendarService(new OAuth2Authenticator<NativeApplicationClient>(provider, GetAuthentication));
            service.Key = "AIzaSyChF6Uf0z9PthPzZ1Qy2lHZ_OrXW6oMUBk";
        }

        private static IAuthorizationState GetAuthentication(NativeApplicationClient arg)
        {
            // Get the auth URL:
            IAuthorizationState state = new AuthorizationState(new[] { CalendarService.Scopes.Calendar.GetStringValue() });
            state.Callback = new Uri(NativeApplicationClient.OutOfBandCallbackUrl);
            state.RefreshToken = Settings.Instance.RefreshToken;
            Uri authUri = arg.RequestUserAuthorization(state);

            IAuthorizationState result = null;

            if (state.RefreshToken == "")
            {
                // Request authorization from the user (by opening a browser window):
                Process.Start(authUri.ToString());

                EnterAuthorizationCode eac = new EnterAuthorizationCode();
                if (eac.ShowDialog() == DialogResult.OK)
                {
                    // Retrieve the access/refresh tokens by using the authorization code:
                    result = arg.ProcessUserAuthorization(eac.authcode, state);

                    //save the refresh token for future use
                    Settings.Instance.RefreshToken = result.RefreshToken;
                    XMLManager.Export(Settings.Instance, MainForm.FILENAME);

                    return result;
                }
                else
                {
                    return null;
                }
            }
            else
            {
                arg.RefreshToken(state, null);
                result = state;
                return result;
            }
        }

        public List<GoogleCalendarListEntry> getCalendars()
        {
            CalendarList request = null;
            request = service.CalendarList.List().Fetch();

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

        public List<Event> getCalendarEntriesInRange(DateTime syncDateTime)
        {
            List<Event> result = new List<Event>();
            Events request = null;

            EventsResource.ListRequest lr = service.Events.List(Settings.Instance.SelectedGoogleCalendar.Id);

            lr.TimeMin = GoogleTimeFrom(syncDateTime.AddDays(-Settings.Instance.DaysInThePast));
            lr.TimeMax = GoogleTimeFrom(syncDateTime.AddDays(+Settings.Instance.DaysInTheFuture + 1));
            lr.SingleEvents = true;
            lr.OrderBy = EventsResource.OrderBy.StartTime;

            request = lr.Fetch();

            // Fetch all pages of events
            if (request != null &&
                request.Items != null)
            {
                result.AddRange(request.Items);
                while (request.NextPageToken != null)
                {
                    lr.PageToken = request.NextPageToken;
                    request = lr.Fetch();
                    result.AddRange(request.Items);
                }
            }

            return result;
        }

        public void deleteCalendarEntry(Event e)
        {
            string request = service.Events.Delete(Settings.Instance.SelectedGoogleCalendar.Id, e.Id).Fetch();
        }

        public void addEntry(Event e)
        {
            var result = service.Events.Insert(e, Settings.Instance.SelectedGoogleCalendar.Id).Fetch();
        }

        //returns the Google Time Format String of a given .Net DateTime value
        //Google Time Format = "2012-08-20T00:00:00+02:00"
        public static string GoogleTimeFrom(DateTime dt)
        {
            string timezone = TimeZoneInfo.Local.GetUtcOffset(dt).ToString();
            if (timezone[0] != '-')
            {
                timezone = '+' + timezone;
            }

            timezone = timezone.Substring(0, 6);
            string result = dt.GetDateTimeFormats('s')[0] + timezone;
            return result;
        }
    }
}
