﻿//----------------------------------------------------------------------------------------------
//    Copyright 2015 Microsoft Corporation
//
//    Licensed under the MIT License (MIT);
//    you may not use this file except in compliance with the License.
//    You may obtain a copy of the License at
//
//      http://mit-license.org/
//
//    Unless required by applicable law or agreed to in writing, software
//    distributed under the License is distributed on an "AS IS" BASIS,
//    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//    See the License for the specific language governing permissions and
//    limitations under the License.
//----------------------------------------------------------------------------------------------

using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.OutlookServices;
using Microsoft.Office365.SharePoint.CoreServices;
using Microsoft.Office365.SharePoint.FileServices;
using O365DataApp.Helpers;
using O365DataApp.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.UI.Popups;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=234238

namespace O365DataApp
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        IDictionary<string, CapabilityDiscoveryResult> AppCapabilities = null;

        public ObservableCollection<MyFile> Files { get; set; }
        public ObservableCollection<MyContact> Contacts { get; set; }
        public ObservableCollection<MyEvent> Events { get; set; }
        public ObservableCollection<MyMail> Mails { get; set; }

        public MainPage()
        {
            Files = new ObservableCollection<MyFile>();
            Contacts = new ObservableCollection<MyContact>();
            Events = new ObservableCollection<MyEvent>();
            Mails = new ObservableCollection<MyMail>();

            this.InitializeComponent();
        }

        private void clearAllLists()
        {
            Files.Clear();
            Contacts.Clear();
            Events.Clear();
            Mails.Clear();
        }

        private void setLoadingString(string msg)
        {
            TextBlock txtBlock = (TextBlock)FindName("loadingTxtBlock");
            txtBlock.Text = msg;
        }

        private async Task getAppCapabilities()
        {
            DiscoveryClient discoveryClient = new DiscoveryClient(
                    async () =>
                    {
                        var authResult = await AuthenticationHelper.GetAccessToken(AuthenticationHelper.DiscoveryServiceResourceId);
                        return authResult.AccessToken;
                    }
                );
            AppCapabilities = await discoveryClient.DiscoverCapabilitiesAsync();
        }

        private async void btnGetMyFiles_Click(object sender, RoutedEventArgs e)
        {
            setLoadingString("Loading..");
            if (AppCapabilities == null)
            {
                await getAppCapabilities();
            }
            clearAllLists();

            var myFilesClient = SpClient.ensureSharepointClientCreated(AppCapabilities, "MyFiles");
            await FileOperations.addFiles(myFilesClient);
            var myFiles = await FileOperations.getMyFiles(myFilesClient);
            foreach (var myFile in myFiles) Files.Add(new MyFile { Name = myFile.Name });

            setLoadingString("");
        }

        private async void btnGetMyContacts_Click(object sender, RoutedEventArgs e)
        {
            setLoadingString("Loading....");
            clearAllLists();
            if (AppCapabilities == null)
            {
                await getAppCapabilities();
            }
            var contactsClient = ExchangeClient.ensureOutlookClientCreated(AppCapabilities, "Contacts");
            string fileAction = FileParse.readFileAction(SettingsHelper.ContactsFilePath);
            if (fileAction == "ADD")
            {
                setLoadingString("Adding Contacts....");
                await ContactsOperations.addContacts(contactsClient, FileParse.readContact());
                var myContacts = await ContactsOperations.getContacts(contactsClient);
                foreach (var myContact in myContacts) Contacts.Add(new MyContact { Name = myContact.DisplayName });
            }
            else if (fileAction == "UPDATE")
            {
                setLoadingString("Updating Contacts....");
                var fileContacts = FileParse.readContact();
                foreach (var fileContact in fileContacts)
                {
                    var contact = await ContactsOperations.getContactByGivenNameAndSurname(contactsClient, fileContact);
                    await ContactsOperations.updateContact(contactsClient, contact.Id, fileContact);
                }
                var myContacts = await ContactsOperations.getContacts(contactsClient);
                foreach (var myContact in myContacts) Contacts.Add(new MyContact { Name = myContact.DisplayName });
            }
            else if (fileAction == "DELETE")
            {
                setLoadingString("Deleting Contacts....");
                var fileContacts = FileParse.readContact();
                foreach (var fileContact in fileContacts)
                {
                    var contact = await ContactsOperations.getContactByGivenNameAndSurname(contactsClient, fileContact);
                    await ContactsOperations.deleteContact(contactsClient, contact.Id);
                }
                var myContacts = await ContactsOperations.getContacts(contactsClient);
                foreach (var myContact in myContacts) Contacts.Add(new MyContact { Name = myContact.DisplayName });
            }
            setLoadingString("");
        }

        private async void btnGetMyEvents_Click(object sender, RoutedEventArgs e)
        {
            setLoadingString("Loading....");
            clearAllLists();
            if (AppCapabilities == null)
            {
                await getAppCapabilities();
            }
            var calendarClient = ExchangeClient.ensureOutlookClientCreated(AppCapabilities, "Calendar");
            string fileAction = FileParse.readFileAction(SettingsHelper.EventsFilePath);
            if (fileAction == "ADD")
            {
                setLoadingString("Adding Events....");
                await CalendarOperations.addEvents(calendarClient, FileParse.readEvents());
                var myEvents = await CalendarOperations.getEvents(calendarClient);
                foreach (var myEvent in myEvents) Events.Add(new MyEvent { Subject = myEvent.Subject });
            }
            else if (fileAction == "UPDATE")
            {
                setLoadingString("Updating Events....");
                var fileEvents = FileParse.readEvents();
                foreach (var fileEvent in fileEvents)
                {
                    var evnt = await CalendarOperations.getEventBySubject(calendarClient, fileEvent);
                    await CalendarOperations.updateEvent(calendarClient, evnt.Id, fileEvent);
                }
                var myEvents = await CalendarOperations.getEvents(calendarClient);
                foreach (var myEvent in myEvents) Events.Add(new MyEvent { Subject = myEvent.Subject });
            }
            else if (fileAction == "DELETE")
            {
                setLoadingString("Deleting Events....");
                var fileEvents = FileParse.readEvents();
                foreach (var fileEvent in fileEvents)
                {
                    var evnt = await CalendarOperations.getEventBySubject(calendarClient, fileEvent);
                    await CalendarOperations.deleteEvent(calendarClient, evnt.Id);
                }
                var myEvents = await CalendarOperations.getEvents(calendarClient);
                foreach (var myEvent in myEvents) Events.Add(new MyEvent { Subject = myEvent.Subject });
            }
            setLoadingString("");
        }

        private async void btnGetMyMails_Click(object sender, RoutedEventArgs e)
        {
            setLoadingString("Loading..");
            clearAllLists();
            if (AppCapabilities == null)
            {
                await getAppCapabilities();
            }

            var mailClient = ExchangeClient.ensureOutlookClientCreated(AppCapabilities, "Mail");
            await MailOperations.sendMail(mailClient, FileParse.readMails());
            var myMails = await MailOperations.getMails(mailClient);
            foreach (var myMail in myMails) Mails.Add(new MyMail { Subject = myMail.Subject });

            setLoadingString("");
        }

        private async void btnLogout_Click(object sender, RoutedEventArgs e)
        {
            await AuthenticationHelper.SignOutAsync();
            clearAllLists();
        }
    }
}
