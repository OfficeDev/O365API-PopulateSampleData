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

using Microsoft.Office365.OutlookServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace O365DataApp.Helpers
{
    static class ContactsOperations
    {
        public static async Task addContacts(OutlookServicesClient contactsClient, List<Contact> newContacts)
        {
            foreach (var newContact in newContacts)
            {
                var refContact = await getContactByGivenNameAndSurname(contactsClient, newContact);
                if (refContact == null)
                {
                    await contactsClient.Me.Contacts.AddContactAsync(newContact);
                }
            }
        }

        public static async Task<List<IContact>> getContacts(OutlookServicesClient contactsClient)
        {
            List<IContact> myContacts = new List<IContact>();
            var contactsResult = await contactsClient.Me.Contacts.OrderBy(c => c.DisplayName).ExecuteAsync();
            do
            {
                var contacts = contactsResult.CurrentPage;
                foreach (var contact in contacts)
                {
                    myContacts.Add(contact);
                }
                contactsResult = await contactsResult.GetNextPageAsync();
            } while (contactsResult != null);
            return myContacts;
        }

        public static async Task<IContact> getContactByGivenNameAndSurname(OutlookServicesClient contactsClient, Contact myContact)
        {
            var contactsResult = await contactsClient.Me.Contacts
                .Where(c => c.GivenName == myContact.GivenName && c.Surname == myContact.Surname)
                .ExecuteSingleAsync();
            return contactsResult;
        }

        public static async Task deleteContact(OutlookServicesClient contactsClient, string contactId)
        {
            var contactToDelete = await contactsClient.Me.Contacts[contactId].ExecuteAsync();
            await contactToDelete.DeleteAsync();
        }

        public static async Task updateContact(OutlookServicesClient contactsClient, string contactId, Contact newContact)
        {
            var contactToUpdate = await contactsClient.Me.Contacts[contactId].ExecuteAsync();
            contactToUpdate.AssistantName = newContact.AssistantName;
            contactToUpdate.BusinessHomePage = newContact.BusinessHomePage;
            contactToUpdate.CompanyName = newContact.CompanyName;
            contactToUpdate.Department = newContact.Department;
            contactToUpdate.DisplayName = newContact.DisplayName;
            contactToUpdate.FileAs = newContact.FileAs;
            contactToUpdate.Generation = newContact.Generation;
            contactToUpdate.GivenName = newContact.GivenName;
            contactToUpdate.Initials = newContact.Initials;
            contactToUpdate.JobTitle = newContact.JobTitle;
            contactToUpdate.Manager = newContact.Manager;
            contactToUpdate.MiddleName = newContact.MiddleName;
            contactToUpdate.MobilePhone1 = newContact.MobilePhone1;
            contactToUpdate.NickName = newContact.NickName;
            contactToUpdate.OfficeLocation = newContact.OfficeLocation;
            contactToUpdate.Profession = newContact.Profession;
            contactToUpdate.Surname = newContact.Surname;
            await contactToUpdate.UpdateAsync();
        }
    }
}
