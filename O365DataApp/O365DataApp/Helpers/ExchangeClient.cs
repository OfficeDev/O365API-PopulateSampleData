using Microsoft.Office365.Discovery;
using Microsoft.Office365.OutlookServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace O365DataApp.Helpers
{
    static class ExchangeClient
    {
        public static OutlookServicesClient ensureOutlookClientCreated(IDictionary<string, CapabilityDiscoveryResult> appCapabilities, string capability)
        {
            var contactsCapability = appCapabilities
                                        .Where(s => s.Key == capability)
                                        .Select(p => new { Key = p.Key, ServiceResourceId = p.Value.ServiceResourceId, ServiceEndPointUri = p.Value.ServiceEndpointUri })
                                        .FirstOrDefault();

            OutlookServicesClient outlookClient = new OutlookServicesClient(contactsCapability.ServiceEndPointUri,
                     async () =>
                     {
                         var authResult = await AuthenticationHelper.GetAccessToken(contactsCapability.ServiceResourceId);
                         return authResult.AccessToken;
                     });

            return outlookClient;
        }
    }
}
