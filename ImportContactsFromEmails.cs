using System.Net;
using System.Text;
using System.Text.Json;
using System.Diagnostics;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace CacheEnterprises.ImportContacts
{
    public class ImportContactsFromEmails
    {
        private readonly ILogger _logger;
        private GraphServiceClient _graphServiceClient;

        public ImportContactsFromEmails(ILoggerFactory loggerFactory)
        {
            _logger = loggerFactory.CreateLogger<ImportContactsFromEmails>();
        }

        [Function("SendEmailAboutMe")]
        public HttpResponseData Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post")] HttpRequestData req)
        {
          var sw = new Stopwatch();
            sw.Restart();

            _logger.LogTrace("CacheEnterprises.ImportContacts - ImportContactsFromEmails.SendEmailAboutMe");
            _logger.LogInformation("Message logged");
    
        //     var domain = "org8.onmicrosoft.com";
        //     var user = "tobyn@cacheenterprises.com.au";
        //     //var pw = "password";
        //     var clientId = "guid-from-portal";
        //     var resource = "https://graph.microsoft.com";
        //     HttpClient client = new HttpClient();

        //     string requestUrl = $"https://login.microsoftonline.com/{domain}/oauth2/token";

           // string request_content = $"grant_type=password&resource={resource}&client_id={clientId}&username={user}&password={pw}&scope=openid+Mail.Send";

            // HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
            // try
            // {
            //     request.Content = new StringContent(request_content, Encoding.UTF8, "application/x-www-form-urlencoded");
            // }
            // catch (Exception x)
            // {
            //     var msg = x.Message;
            // }
            // HttpResponseMessage cliResp = client.Send(request);

            // string responseString = cliResp.ToString();
            // _logger.LogInformation(responseString);
            // GenericToken token = JsonSerializer.Deserialize<GenericToken>(responseString);
            // var at = token.access_token;

            // var me = GetUserInfo(at);

            // _logger.LogInformation($"Display Name:{me.displayName}\nUpn:{me.userPrincipalName}\nPreferred Language:{me.preferredLanguage}");

            // SendMail(at);

            //create response 
            var response = req.CreateResponse(HttpStatusCode.OK);

            response.Headers.Add("Date", "Mon, 18 Jul 2016 16:06:00 GMT");
            response.Headers.Add("Content-Type", "text/html; charset=utf-8");
            response.WriteString("Email Sent.");

            _logger.LogTrace(string.Format(@"funcExecutionTimeMs {0}", sw.Elapsed.TotalMilliseconds,
                new Dictionary<string, object> {
                    { "foo", "bar" },
                    { "baz", 42 }
                })
            );
            sw.Stop();
            return response;
        }
    
        private MailMessage SendMail(string token)
        {
            //https://graph.microsoft.io/en-us/docs/api-reference/v1.0/api/user_sendmail
            HttpClient client = new HttpClient();
            var authHeader = "Bearer " + token;
            string requestUrl = $"https://graph.microsoft.com/v1.0/me/sendMail";

            var baseEmail = new EmailAddress { name = "Alice", address = "alice@contoso.com" };
            var fromEmail = new EmailAddress { name = "Bob", address = "bob@contoso.com" };
            var from = new From { emailAddress = fromEmail };
            var send = new Sender { emailAddress = fromEmail };

            var recList = new List<ToRecipient>();
            recList.Add(new ToRecipient { emailAddress = baseEmail });

            var replyAddress = new List<ReplyTo>();
            replyAddress.Add(new ReplyTo { emailAddress = fromEmail });

            var bcc = new List<BccRecipient>();
            var cc = new List<CcRecipient>();
            var cats = new List<string>();

            var body = new Body { contentType = "text", content = "Hello World" };

            var draft = new MailItem { toRecipients = recList,sender=send,from = from, bccRecipients=bcc,ccRecipients=cc, replyTo=replyAddress, categories=cats, body = body, subject = "Hello World" };

            var Message = new MailMessage { Message = draft, SaveToSentItems = "true" };

            string request_content = JsonSerializer.Serialize(Message);

            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
            try
            {
                client.DefaultRequestHeaders.TryAddWithoutValidation("Authorization", authHeader);
                request.Content = new StringContent(request_content, Encoding.UTF8, "application/json");
            }
            catch (Exception x)
            {
                var exception = x.Message;
            }
            HttpResponseMessage response = client.Send(request);

            string responseString = response.Content.ToString();
            _logger.LogTrace(responseString);
            MailMessage msg = JsonSerializer.Deserialize<MailMessage>(json: responseString);
            return msg;
        }

        private AADUser GetUserInfo(string token)
        {
            string graphRequest = $"https://graph.microsoft.com/v1.0/me/";
            var authHeader = "Bearer " + token;
            HttpClient client = new HttpClient();

            client.DefaultRequestHeaders.TryAddWithoutValidation("Authorization", authHeader);
            var response = client.GetAsync(new Uri(graphRequest));

            string content = response.ToString();
            var user = JsonSerializer.Deserialize<AADUser>(content);
            return user;
        }
    }

public class GenericToken
{
    public string token_type { get; set; }
    public string scope { get; set; }
    public string resource { get; set; }
    public string access_token { get; set; }
    public string refresh_token { get; set; }
    public string id_token { get; set; }
    public string expires_in { get; set; }
}

public class AADUser
{
    public string displayName { get; set; }
    public string givenName { get; set; }
    public string surname { get; set; }
    public string mail { get; set; }
    public string preferredLanguage { get; set; }
    public string userPrincipalName { get; set; }
    public string mobilePhone { get; set; }
}

public class MailMessage
{
    public MailItem Message { get; set; }
    public string SaveToSentItems { get; set; }
}

public class EmailAddress
{
    public string name { get; set; }
    public string address { get; set; }
}

public class BccRecipient
{
    public EmailAddress emailAddress { get; set; }
}

public class Body
{
    public string contentType { get; set; }
    public string content { get; set; }
}

public class CcRecipient
{
    public EmailAddress emailAddress { get; set; }
}

public class From
{
    public EmailAddress emailAddress { get; set; }
}

public class ReplyTo
{
    public EmailAddress emailAddress { get; set; }
}

public class Sender
{
    public EmailAddress emailAddress { get; set; }
}

public class ToRecipient
{
    public EmailAddress emailAddress { get; set; }
}

public class UniqueBody
{
    public string contentType { get; set; }
    public string content { get; set; }
}

public class MailItem
{
    public List<BccRecipient> bccRecipients { get; set; }
    public Body body { get; set; }
    public string bodyPreview { get; set; }
    public List<string> categories { get; set; }
    public List<CcRecipient> ccRecipients { get; set; }
    public string changeKey { get; set; }
    public string conversationId { get; set; }
    public string createdDateTime { get; set; }
    public From from { get; set; }
    public bool hasAttachments { get; set; }
    public string id { get; set; }
    public string importance { get; set; }
    public string inferenceClassification { get; set; }
    public string internetMessageId { get; set; }
    public bool isDeliveryReceiptRequested { get; set; }
    public bool isDraft { get; set; }
    public bool isRead { get; set; }
    public bool isReadReceiptRequested { get; set; }
    public string lastModifiedDateTime { get; set; }
    public string parentFolderId { get; set; }
    public string receivedDateTime { get; set; }
    public List<ReplyTo> replyTo { get; set; }
    public Sender sender { get; set; }
    public string sentDateTime { get; set; }
    public string subject { get; set; }
    public List<ToRecipient> toRecipients { get; set; }
    public UniqueBody uniqueBody { get; set; }
    public string webLink { get; set; }
}
}
