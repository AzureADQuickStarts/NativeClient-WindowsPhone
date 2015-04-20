using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.ApplicationModel.Activation;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Popups;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=391641

namespace DirectorySearcher
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page, IWebAuthenticationContinuable
    {

        #region Config Values

        //
        // The Client ID is used by the application to uniquely identify itself to Azure AD.
        // The Tenant is the name of the Azure AD tenant in which this application is registered.
        // The AAD Instance is the instance of Azure, for example public Azure or Azure China.
        // The Authority is the sign-in URL of the tenant.
        //

        const string aadInstance = "https://login.microsoftonline.com/{0}";
        const string tenant = "[Enter tenant name, e.g. contoso.onmicrosoft.com]";
        const string clientId = "[Enter client ID as obtained from Azure Portal, e.g. 82692da5-a86f-44c9-9d53-2f88d52b478b]";

        static string authority = String.Format(CultureInfo.InvariantCulture, aadInstance, tenant);

        const string graphResourceId = "https://graph.windows.net/";
        const string graphEndpoint = "https://graph.windows.net/";
        const string graphApiVersion = "1.5";

        #endregion

        private HttpClient httpClient = new HttpClient();
        private AuthenticationContext authContext = null;
        private Uri redirectURI = null;

        public MainPage()
        {
            this.InitializeComponent();

            this.NavigationCacheMode = NavigationCacheMode.Required;

            //
            // Every Windows Store application has a unique URI.
            // Windows ensures that only this application will receive messages sent to this URI.
            // ADAL uses this URI as the application's redirect URI to receive OAuth responses.
            // 
            // To determine this application's redirect URI, which is necessary when registering the app
            //      in AAD, set a breakpoint on the next line, run the app, and copy the string value of the URI.
            //      This is the only purposes of this line of code, it has no functional purpose in the application.
            //
            redirectURI = Windows.Security.Authentication.Web.WebAuthenticationBroker.GetCurrentApplicationCallbackUri();

            // TODO: Initialize the AuthenticationContext
        }

        #region IWebAuthenticationContinuable implementation

        // TODO: Implement the ContinueWebAuthenticationInterface
        
        #endregion

        /// <summary>
        /// Invoked when this page is about to be displayed in a Frame.
        /// </summary>
        /// <param name="e">Event data that describes how this page was reached.
        /// This parameter is typically used to configure the page.</param>
        protected override void OnNavigatedTo(NavigationEventArgs e)
        {

        }

        #region Button Handlers

        // clear the token cache
        private void RemoveAppBarButton_Click(object sender, RoutedEventArgs e)
        {
            SignOut();
        }

        private void SignOut()
        {
            // TODO: Clear the token cache of all tokens

            // Reset UI elements
            SearchResults.ItemsSource = null;
            SearchTermText.Text = string.Empty;
            ActiveUser.Text = string.Empty;
            StatusResult.Text = string.Empty;
        }


        private async void Search(object sender, RoutedEventArgs e)
        {
            if (SearchTermText.Text.Length <= 0)
            {
                MessageDialog dialog = new MessageDialog("Please enter a valid search term.");
                await dialog.ShowAsync();
                return;
            }

            // TODO: Acquire a token to call the Graph API
        }

        #endregion

        private async void QueryGraph(AuthenticationResult result)
        {

            // TODO: Attach the access token to the HTTP request

            // TODO: Display the user's Id on the main page.

            string graphRequest = String.Format(CultureInfo.InvariantCulture, "{0}{1}/users?api-version={2}&$filter=startswith(userPrincipalName, '{3}')", graphEndpoint, tenant, graphApiVersion, SearchTermText.Text);
            HttpResponseMessage response = await httpClient.GetAsync(graphRequest);

            if (response.StatusCode == HttpStatusCode.Unauthorized)
            {
                // If the Graph API returned a 401, user may need to sign in again.
                MessageDialog dialog = new MessageDialog("Sorry, you don't have access to the Graph API.  Please sign-in again.");
                await dialog.ShowAsync();
                SignOut();
                return;
            }

            if (!response.IsSuccessStatusCode)
            {
                MessageDialog dialog = new MessageDialog("Sorry, an error occurred accessing the Graph API.  Please try again.");
                await dialog.ShowAsync();
                return;
            }

            string content = response.Content.ReadAsStringAsync().Result;
            JObject jResult = JObject.Parse(content);

            if (jResult["odata.error"] != null)
            {
                MessageDialog dialog = new MessageDialog("Sorry, an error occurred accessing the Graph API.  Please try again.");
                await dialog.ShowAsync();
                return;
            }

            if (jResult["value"].Count() <= 0)
            {
                StatusResult.Text = "No Users Found";
                StatusResult.Foreground = new SolidColorBrush(Windows.UI.Colors.Red);
            }

            StatusResult.Text = "Success";
            StatusResult.Foreground = new SolidColorBrush(Windows.UI.Colors.Green);
            SearchResults.ItemsSource = 
                from user in jResult["value"]
                select new
                {
                    userPrincipalName = user["userPrincipalName"],
                    displayName = user["displayName"]
                };
        }
    }
}
