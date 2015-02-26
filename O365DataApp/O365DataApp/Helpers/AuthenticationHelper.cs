using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Threading.Tasks;
using Windows.Security.Authentication.Web;
using Microsoft.Office365.OAuth;
using Windows.Storage;
using Windows.UI.Popups;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.Office365.OutlookServices;
using Microsoft.Office365.SharePoint.CoreServices;
using Microsoft.Office365.Discovery;

namespace O365DataApp.Helpers
{
    public partial class AuthenticationHelper
    {
        public static readonly string DiscoveryServiceResourceId = "https://api.office.com/discovery/";
        const string AuthorityFormat = "https://login.windows.net/{0}/";
        static readonly Uri DiscoveryServiceEndpointUri = new Uri("https://api.office.com/discovery/v1.0/me/");
        static readonly string ClientId = App.Current.Resources["ida:ClientID"].ToString();
        static ApplicationDataContainer AppSettings = ApplicationData.Current.LocalSettings;
        static string _authority = String.Empty;
        static string _lastTenantId = "common";
        const string _lastTenantIdKey = "LastAuthority";
        static AuthenticationContext authContext = null;

        public static Uri AppRedirectURI
        {
            get
            {
                return WebAuthenticationBroker.GetCurrentApplicationCallbackUri();
            }
        }

        public static string LastTenantId
        {
            get
            {
                if (AppSettings.Values.ContainsKey(_lastTenantIdKey) && AppSettings.Values[_lastTenantIdKey] != null)
                {
                    return AppSettings.Values[_lastTenantIdKey].ToString();
                }
                else
                {
                    return _lastTenantId;
                }

            }

            set
            {
                _lastTenantId = value;
                AppSettings.Values[_lastTenantIdKey] = _lastTenantId;
            }
        }

        public static string Authority
        {
            get
            {
                _authority = String.Format(AuthorityFormat, LastTenantId);

                return _authority;
            }
        }

        public static async Task<AuthenticationResult> GetAccessToken(string serviceResourceId)
        {
            AuthenticationResult authResult = null;

            if (authContext == null)
            {
                authContext = new AuthenticationContext(Authority);

                #region To enable Windows Integrated Authentication (if you deploying your app in a corporate network)

                // To enable Windows Integrated Authentication, in Package.appxmanifest, in the Capabilities tab, enable:
                // * Enterprise Authentication
                // * Private Networks (Client & Server)
                // * Shared User Certificates
                // Plus add the following line of code:
                // 
                //authContext.UseCorporateNetwork = true;

                #endregion

                authResult = await authContext.AcquireTokenAsync(serviceResourceId, ClientId, AppRedirectURI);
            }
            else
            {
                authResult = await authContext.AcquireTokenSilentAsync(serviceResourceId, ClientId);
            }

            LastTenantId = authResult.TenantId;

            if (authResult.Status != AuthenticationStatus.Success)
            {
                LastTenantId = authResult.TenantId;

                if (authResult.Error == "authentication_canceled")
                {
                    // The user cancelled the sign-in, no need to display a message.
                }
                else
                {
                    MessageDialog dialog = new MessageDialog(string.Format("If the error continues, please contact your administrator.\n\nError: {0}\n\n Error Description:\n\n{1}", authResult.Error, authResult.ErrorDescription), "Sorry, an error occurred while signing you in.");
                    await dialog.ShowAsync();
                }
            }

            return authResult;
        }

        public static async Task SignOutAsync()
        {
            //await authContext.LogoutAsync(LastTenantId);
            //authContext.TokenCache.Clear();
            ApplicationData.Current.LocalSettings.Values["TenantId"] = null;
            ApplicationData.Current.LocalSettings.Values["LastAuthority"] = null;
        }

    }
}
