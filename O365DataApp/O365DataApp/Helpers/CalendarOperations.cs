using Microsoft.Office365.OutlookServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace O365DataApp.Helpers
{
    static class CalendarOperations
    {
        public static async Task addEvents(OutlookServicesClient calendarClient, List<Event> newEvents)
        {
            foreach (var newEvent in newEvents)
            {
                var refEvent = await getEventBySubject(calendarClient, newEvent);
                if (refEvent == null)
                {
                    await calendarClient.Me.Events.AddEventAsync(newEvent);
                }
            }
        }

        public static async Task<List<Event>> getEvents(OutlookServicesClient calendarClient)
        {
            List<Event> myEvents = new List<Event>();
            var eventsResult = await calendarClient.Me.Events.OrderBy(c => c.Start).ExecuteAsync();
            do
            {
                var events = eventsResult.CurrentPage;
                foreach (var evnt in events)
                {
                    myEvents.Add((Event)evnt);
                }
                eventsResult = await eventsResult.GetNextPageAsync();
            } while (eventsResult != null);
            return myEvents;
        }

        public static async Task<IEvent> getEventBySubject(OutlookServicesClient calendarClient, Event refEvent)
        {
            var evnt = await calendarClient.Me.Events.Where(c => c.Subject == refEvent.Subject).ExecuteSingleAsync();
            return evnt;
        }

        public static async Task deleteEvent(OutlookServicesClient calendarClient, string eventId)
        {
            var eventToDelete = await calendarClient.Me.Calendar.Events.GetById(eventId).ExecuteAsync();
            await eventToDelete.DeleteAsync(false);
        }

        public static async Task updateEvent(OutlookServicesClient calendarClient, string eventId, Event newEvent)
        {
            var eventToUpdate = await calendarClient.Me.Calendar.Events.GetById(eventId).ExecuteAsync();
            eventToUpdate.Body = newEvent.Body;
            eventToUpdate.BodyPreview = newEvent.BodyPreview;
            eventToUpdate.Location = newEvent.Location;
            eventToUpdate.Organizer = newEvent.Organizer;
            eventToUpdate.Subject = newEvent.Subject;
            eventToUpdate.Attendees.Clear();
            foreach (var attendee in newEvent.Attendees)
            {
                eventToUpdate.Attendees.Add(attendee);
            }
            await eventToUpdate.UpdateAsync();
        }
    }
}
