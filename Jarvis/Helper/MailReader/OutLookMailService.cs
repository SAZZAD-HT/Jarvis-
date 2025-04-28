using System.Net.Http.Headers;
using System.Text.Json;
using System.Text;
using Microsoft.Extensions.Configuration;

namespace Jarvis.Helper.MailReader
{
    public class OutLookMailService
    {
        private readonly HttpClient _httpClient;
        private readonly IConfiguration _configuration;
        private string _accessToken;

        public OutLookMailService(IConfiguration configuration)
        {
            _httpClient = new HttpClient();
            _configuration = configuration;
        }

        private async Task AuthenticateAsync()
        {
            var tenantId = _configuration["AzureAd:TenantId"];
            var clientId = _configuration["AzureAd:ClientId"];
            var clientSecret = _configuration["AzureAd:ClientSecret"];

            var url = $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token";

            var body = new Dictionary<string, string>
            {
                { "client_id", clientId },
                { "scope", "https://graph.microsoft.com/.default" },
                { "client_secret", clientSecret },
                { "grant_type", "client_credentials" }
            };

            var response = await _httpClient.PostAsync(url, new FormUrlEncodedContent(body));
            var json = await response.Content.ReadAsStringAsync();
            var result = JsonDocument.Parse(json);

            _accessToken = result.RootElement.GetProperty("access_token").GetString();
        }

        public async Task<List<OutlookMessage>> GetUnreadOrderMailsFromRoboFolderAsync(string userEmail)
        {
            if (string.IsNullOrEmpty(_accessToken))
                await AuthenticateAsync();

            var requestUrl = $"https://graph.microsoft.com/v1.0/users/{userEmail}/mailFolders/Robo/messages" +
                             "?$filter=isRead eq false and startsWith(subject, 'Order:')" +
                             "&$select=id,subject,receivedDateTime,bodyPreview,from" +
                             "&$top=10" +
                             "&$orderby=receivedDateTime desc";

            var request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken);

            var response = await _httpClient.SendAsync(request);
            response.EnsureSuccessStatusCode();

            var json = await response.Content.ReadAsStringAsync();
            var mailResult = JsonDocument.Parse(json);

            var messages = new List<OutlookMessage>();

            if (mailResult.RootElement.TryGetProperty("value", out JsonElement mails))
            {
                foreach (var mail in mails.EnumerateArray())
                {
                    messages.Add(new OutlookMessage
                    {
                        Id = mail.GetProperty("id").GetString(),
                        Subject = mail.GetProperty("subject").GetString(),
                        ReceivedDateTime = mail.GetProperty("receivedDateTime").GetDateTime(),
                        BodyPreview = mail.GetProperty("bodyPreview").GetString(),
                        From = mail.GetProperty("from").GetProperty("emailAddress").GetProperty("address").GetString()
                    });
                }
            }

            return messages;
        }

        public async Task SendEmailAsync(string recipient, string subject, string body)
        {
            if (string.IsNullOrEmpty(_accessToken))
                await AuthenticateAsync();

            var message = new
            {
                message = new
                {
                    subject = subject,
                    body = new
                    {
                        contentType = "Text",
                        content = body
                    },
                    toRecipients = new[]
                    {
                        new
                        {
                            emailAddress = new
                            {
                                address = recipient
                            }
                        }
                    }
                },
                saveToSentItems = "false"
            };

            var requestContent = new StringContent(JsonSerializer.Serialize(message), Encoding.UTF8, "application/json");

            var request = new HttpRequestMessage(HttpMethod.Post, $"https://graph.microsoft.com/v1.0/users/{recipient}/sendMail")
            {
                Content = requestContent
            };
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken);

            var response = await _httpClient.SendAsync(request);
            response.EnsureSuccessStatusCode();
        }
    }

    public class OutlookMessage
    {
        public string Id { get; set; }
        public string Subject { get; set; }
        public DateTime? ReceivedDateTime { get; set; }
        public string BodyPreview { get; set; }
        public string From { get; set; }
    }
}
