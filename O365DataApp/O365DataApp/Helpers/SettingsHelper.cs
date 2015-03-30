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
