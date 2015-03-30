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

using Microsoft.Office365.Discovery;
using Microsoft.Office365.SharePoint.CoreServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace O365DataApp.Helpers
{
    static class SpClient
    {
        public static SharePointClient ensureSharepointClientCreated(IDictionary<string, CapabilityDiscoveryResult> appCapabilities, string capability)
        {
            var contactsCapability = appCapabilities
                                        .Where(s => s.Key == capability)
                                        .Select(p => new { Key = p.Key, ServiceResourceId = p.Value.ServiceResourceId, ServiceEndPointUri = p.Value.ServiceEndpointUri })
                                        .FirstOrDefault();

            SharePointClient spClient = new SharePointClient(contactsCapability.ServiceEndPointUri,
                     async () =>
                     {
                         var authResult = await AuthenticationHelper.GetAccessToken(contactsCapability.ServiceResourceId);
                         return authResult.AccessToken;
                     });

            return spClient;
        }
    }
}
