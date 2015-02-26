using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.ApplicationModel;

namespace O365DataApp.Helpers
{
    class SettingsHelper
    {
        private static string _contactsFilePath = Path.Combine(Package.Current.InstalledLocation.Path, "Assets/AddContact.xml");
        private static string _eventsFilePath = Path.Combine(Package.Current.InstalledLocation.Path, "Assets/AddEvent.xml");
        private static string _mailsFilePath = Path.Combine(Package.Current.InstalledLocation.Path, "Assets/AddMail.xml");

        public static string ContactsFilePath
        {
            get
            {
                return _contactsFilePath;
            }
        }

        public static string EventsFilePath
        {
            get
            {
                return _eventsFilePath;
            }
        }

        public static string MailsFilePath
        {
            get
            {
                return _mailsFilePath;
            }
        }
    }
}
