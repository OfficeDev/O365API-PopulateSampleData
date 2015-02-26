using Microsoft.Office365.OutlookServices;
using Microsoft.Office365.SharePoint.FileServices;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Windows.ApplicationModel;
using Windows.Storage;
using Windows.UI.Popups;

namespace O365DataApp.Helpers
{
    public static class FileParse
    {
        public static string readFileAction(string filePath)
        {
            string result = null;
            XDocument loadedData = XDocument.Load(filePath);
            result = loadedData.Descendants("Action").Elements("ActionName").FirstOrDefault().Value;
            return result;
        }

        public static List<Contact> readContact()
        {
            XDocument loadedData = XDocument.Load(SettingsHelper.ContactsFilePath);
            List<Contact> data = (from query in loadedData.Descendants("Contact")
                                  select new Contact
                                  {
                                      AssistantName = (string)query.Element("AssistantName") != "" ? (string)query.Element("AssistantName") : null,
                                      BusinessHomePage = (string)query.Element("BusinessHomePage") != "" ? (string)query.Element("BusinessHomePage") : null,
                                      //BusinessPhones = ,
                                      CompanyName = (string)query.Element("CompanyName") != "" ? (string)query.Element("CompanyName") : null,
                                      Department = (string)query.Element("Department") != "" ? (string)query.Element("Department") : null,
                                      DisplayName = (string)query.Element("DisplayName") != "" ? (string)query.Element("DisplayName") : null,
                                      EmailAddresses = (from newQuery in query.Descendants("EmailAddress")
                                                        select new EmailAddress
                                                        {
                                                            Name = (string)newQuery.Value,
                                                            Address = (string)newQuery.Value
                                                        }).ToList<EmailAddress>(),
                                      FileAs = (string)query.Element("FileAs") != "" ? (string)query.Element("FileAs") : null,
                                      Generation = (string)query.Element("Generation") != "" ? (string)query.Element("Generation") : null,
                                      GivenName = (string)query.Element("GivenName") != "" ? (string)query.Element("GivenName") : null,
                                      //HomePhones =,
                                      //ImAddresses = ,
                                      Initials = (string)query.Element("Initials") != "" ? (string)query.Element("Initials") : null,
                                      JobTitle = (string)query.Element("JobTitle") != "" ? (string)query.Element("JobTitle") : null,
                                      Manager = (string)query.Element("Manager") != "" ? (string)query.Element("Manager") : null,
                                      MiddleName = (string)query.Element("MiddleName") != "" ? (string)query.Element("MiddleName") : null,
                                      MobilePhone1 = (string)query.Element("MobilePhone1") != "" ? (string)query.Element("MobilePhone1") : null,
                                      NickName = (string)query.Element("NickName") != "" ? (string)query.Element("NickName") : null,
                                      OfficeLocation = (string)query.Element("OfficeLocation") != "" ? (string)query.Element("OfficeLocation") : null,
                                      Profession = (string)query.Element("Profession") != "" ? (string)query.Element("Profession") : null,
                                      Surname = (string)query.Element("Surname") != "" ? (string)query.Element("Surname") : null
                                  }).ToList<Contact>();
            return data;
        }

        public static List<Event> readEvents()
        {
            XDocument loadedData = XDocument.Load(SettingsHelper.EventsFilePath);
            List<Event> data = (from query in loadedData.Descendants("Event")
                                select new Event
                                {
                                    Attendees = (from newQuery in query.Descendants("Attendee")
                                                 select new Attendee
                                                 {
                                                     EmailAddress = new EmailAddress() { Name = (string)newQuery.Value, Address = (string)newQuery.Value },
                                                     Type = AttendeeType.Required
                                                 }).ToList<Attendee>(),
                                    Body = (from newQuery in query.Descendants("Body")
                                            select new ItemBody
                                            {
                                                Content = newQuery.Value,
                                                ContentType = BodyType.Text
                                            }).FirstOrDefault(),
                                    BodyPreview = (string)query.Element("BodyPreview"),
                                    End = DateTimeOffset.Now.AddMinutes(30),
                                    Location = (from newQuery in query.Descendants("Location")
                                                select new Location
                                                {
                                                    DisplayName = (string)newQuery.Value
                                                }).FirstOrDefault(),
                                    Organizer = (from newQuery in query.Descendants("Organizer")
                                                 select new Recipient
                                                 {
                                                     EmailAddress = new EmailAddress() { Name = (string)newQuery.Value, Address = (string)newQuery.Value }
                                                 }).FirstOrDefault(),
                                    Start = DateTimeOffset.Now,
                                    Subject = (string)query.Element("Subject"),
                                }).ToList<Event>();
            return data;
        }

        public static List<Message> readMails()
        {
            XDocument loadedData = XDocument.Load(SettingsHelper.MailsFilePath);

            List<Message> data = (from query in loadedData.Descendants("Message")
                                  select new Message
                                  {
                                      BccRecipients = (from newQuery in query.Descendants("BccRecipient")
                                                       select new Recipient
                                                       {
                                                           EmailAddress = new EmailAddress() { Name = (string)newQuery.Value, Address = (string)newQuery.Value }
                                                       }).ToList<Recipient>(),
                                      Body = (from newQuery in query.Descendants("Body")
                                              select new ItemBody
                                              {
                                                  Content = newQuery.Value,
                                                  ContentType = BodyType.Text
                                              }).FirstOrDefault(),
                                      BodyPreview = (string)query.Element("BodyPreview"),
                                      CcRecipients = (from newQuery in query.Descendants("CcRecipient")
                                                      select new Recipient
                                                      {
                                                          EmailAddress = new EmailAddress() { Name = (string)newQuery.Value, Address = (string)newQuery.Value }
                                                      }).ToList<Recipient>(),
                                      Subject = (string)query.Element("Subject"),
                                      ToRecipients = (from newQuery in query.Descendants("ToRecipient")
                                                      select new Recipient
                                                      {
                                                          EmailAddress = new EmailAddress() { Name = (string)newQuery.Value, Address = (string)newQuery.Value }
                                                      }).ToList<Recipient>()
                                  }).ToList<Message>();
            return data;
        }
    }
}
