using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Diagnostics;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;

namespace CacheEnterprises.ImportContacts
{
    public class GetUserMicrosoftGraph
    {
        private GraphServiceClient _graphServiceClient;

        [Function("GetUserMicrosoftGraph")]
        public HttpResponseData Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post")] HttpRequestData req)
        {
           GraphServiceClient graphClient = new GraphServiceClient( CreateAuthorizationProvider() );

                var user = graphClient.Me.Messages
                .Request()
                .Filter("importance eq 'high'")
                .GetAsync();

          //  var serializedUser = JsonSerializer.Serialize(user);  
            //create response 
            var response = req.CreateResponse(HttpStatusCode.OK);

            response.Headers.Add("Date", "Mon, 18 Jul 2016 16:06:00 GMT");
            response.Headers.Add("Content-Type", "text/html; charset=utf-8");
            response.WriteAsJsonAsync(user);
            response.WriteString("Email Sent.");

            return response;

        }

        private GraphServiceClient GetAuthenticatedGraphClient()
        {
            var authenticationProvider = CreateAuthorizationProvider();
            _graphServiceClient = new GraphServiceClient(authenticationProvider);
            return _graphServiceClient;
        }
        private IAuthenticationProvider CreateAuthorizationProvider()
        {
            var clientId = System.Environment.GetEnvironmentVariable("AzureADAppClientId", EnvironmentVariableTarget.Process);
            var clientSecret = System.Environment.GetEnvironmentVariable("AzureADAppClientSecret", EnvironmentVariableTarget.Process);
            var redirectUri = System.Environment.GetEnvironmentVariable("AzureADAppRedirectUri", EnvironmentVariableTarget.Process);
            var tenantId = System.Environment.GetEnvironmentVariable("AzureADAppTenantId", EnvironmentVariableTarget.Process);
            var authority = $"https://login.microsoftonline.com/{tenantId}/v2.0";

            //this specific scope means that application will default to what is defined in the application registration rather than using dynamic scopes
            List<string> scopes = new List<string>();
            scopes.Add("https://graph.microsoft.com/.default");

            var cca = ConfidentialClientApplicationBuilder.Create(clientId)
                                                .WithAuthority(authority)
                                                .WithRedirectUri(redirectUri)
                                                .WithClientSecret(clientSecret)
                                                .Build();

            return new MsalAuthenticationProvider(cca, scopes.ToArray());;
        }
    }

}