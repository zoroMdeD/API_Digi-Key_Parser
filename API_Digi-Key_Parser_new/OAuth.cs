using ApiClient.Extensions;
using ApiClient.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace API_Digi_Key_Parser_new
{
    public class OAuth
    {
        private ApiClientSettings _clientSettings;

        public OAuth()
        {

        }

        public async Task<string> Authorize()
        {
            try
            {
                // read clientSettings values from apiclient.config
                _clientSettings = ApiClientSettings.CreateFromConfigFile();
                //Console.WriteLine(_clientSettings.ToString());

                // start up a HttpListener for the callback(RedirectUri) from the OAuth2 server
                var httpListener = new HttpListener();
                httpListener.Prefixes.Add(_clientSettings.RedirectUri.EnsureTrailingSlash());
                //Console.WriteLine($"listening to {_clientSettings.RedirectUri}");
                httpListener.Start();

                // Initialize our OAuth2 service
                var oAuth2Service = new ApiClient.OAuth2.OAuth2Service(_clientSettings);
                var scopes = "";

                // create Authorize url and send call it thru Process.Start
                var authUrl = oAuth2Service.GenerateAuthUrl(scopes);
                Process.Start(authUrl);

                // get the URL returned from the callback(RedirectUri)
                var context = await httpListener.GetContextAsync();

                // Done with the callback, so stop the HttpListener
                httpListener.Stop();

                // exact the query parameters from the returned URL
                var queryString = context.Request.Url.Query;
                var queryColl = HttpUtility.ParseQueryString(queryString);

                // Grab the needed query parameter code from the query collection
                var code = queryColl["code"];

                // Pass the returned code value to finish the OAuth2 authorization
                var result = await oAuth2Service.FinishAuthorization(code);

                // Check if you got an error during finishing the OAuth2 authorization
                if (result.IsError)
                {
                    return $"\n\nError            : {result.Error}" + Environment.NewLine +
                           $"\n\nError.Description: {result.ErrorDescription}";
                }
                else
                {
                    _clientSettings.UpdateAndSave(result);

                    return _clientSettings.ToString() + Environment.NewLine +
                           $"listening to {_clientSettings.RedirectUri}" + Environment.NewLine +
                           $"Using code {code}" + Environment.NewLine +
                           $"Access token : {result.AccessToken}" + Environment.NewLine +
                           $"Refresh token: {result.RefreshToken}" + Environment.NewLine +
                           $"Expires in   : {result.ExpiresIn}" + Environment.NewLine + Environment.NewLine +
                            "After a good refresh" + Environment.NewLine;
                }
            }
            catch(Exception)
            {
                throw;
            }
        }
    }
}
