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
