using Microsoft.Graph.API.Demo.Models;
using System.Net.Http.Headers;
using System.Text;

namespace Microsoft.Graph.API.Demo
{
    class Program
    {
        static async Task Main(string[] args)
        {
            while (true)
            {
                Console.Clear();
                string clientId = "2312ce3c-1c8b-4f31-9b28-b69420deb853";
                string clientSecret = "CLY8Q~O0TEXyPY-kAHEtcHfV97J9G3K4Krw5vap4";
                string tenantId = "112c896a-ba88-461d-a9a9-80c4d2aca596";
                string senderEmail = "RViramgama@ProlificsPOC.com";
                string recipientEmail = "priyanka.harne@prolifics.com";

                Console.Write("Enter choice \n\t 1. Send Email \n\t 2. Read Email \n\t Enter your choice: ");
                string value = Console.ReadLine();

                string accessToken = await GetAccessToken(clientId, clientSecret, tenantId);

                if (!string.IsNullOrEmpty(accessToken))
                {
                    switch (value)
                    {
                        case "1":
                            SendEmail(accessToken, senderEmail, recipientEmail);
                            break;
                        default:
                            var messages = await GetEmails(accessToken, senderEmail);
                            PrintMessages(messages);
                            break;
                    }
                }
                else
                {
                    Console.WriteLine("Failed to acquire access token.");
                }
                Console.ReadKey();
            }
        }

        static void PrintMessages(Messages messages)
        {
            foreach (var message in messages.value)
            {
                Console.WriteLine($"Sender: {message.sender.emailAddress.address} Receiver: {string.Join(";", message.toRecipients.Select(x => x.emailAddress.address))} - Subject: {message.subject}");
            }
        }

        static async Task<Messages> GetEmails(string token, string userEmail)
        {
            try
            {
                string endpoint = $"https://graph.microsoft.com/v1.0/users/{userEmail}/messages";

                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    HttpResponseMessage response = await client.GetAsync(endpoint);

                    if (response.IsSuccessStatusCode)
                    {
                        string jsonResponse = await response.Content.ReadAsStringAsync();
                        return Newtonsoft.Json.JsonConvert.DeserializeObject<Messages>(jsonResponse);
                    }
                    else
                    {
                        Console.WriteLine($"Failed to send email. Status code: {response.StatusCode}");
                        throw new Exception(response.StatusCode.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                throw;
            }
        }

        static async Task SendEmail(string token, string fromEmail, string toEmail)
        {
            try
            {
                string endpoint = $"https://graph.microsoft.com/v1.0/users/{fromEmail}/sendMail";

                string jsonBody = $@"
                    {{
                        ""message"":
                        {{
                            ""subject"": ""Test Email"",
                            ""body"":
                            {{
                                ""contentType"": ""Text"",
                                ""content"": ""This is a test email sent via Microsoft Graph API.""
                            }},
                            ""toRecipients"": [
                                {{
                                    ""emailAddress"": {{
                                        ""address"": ""{toEmail}""
                                    }}
                                }}
                            ]
                        }}
                    }}";

                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    HttpResponseMessage response = await client.PostAsync(endpoint, new StringContent(jsonBody, Encoding.UTF8, "application/json"));

                    if (response.IsSuccessStatusCode)
                    {
                        Console.WriteLine("Email sent successfully.");
                    }
                    else
                    {
                        Console.WriteLine($"Failed to send email. Status code: {response.StatusCode}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        static async Task<string> GetAccessToken(string clientId, string clientSecret, string tenantId)
        {
            string authority = $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token";

            using (HttpClient client = new HttpClient())
            {
                var requestContent = new FormUrlEncodedContent(new[]
                {
            new KeyValuePair<string, string>("client_id", clientId),
            new KeyValuePair<string, string>("client_secret", clientSecret),
            new KeyValuePair<string, string>("scope", "https://graph.microsoft.com/.default"),
            new KeyValuePair<string, string>("grant_type", "client_credentials")
        });

                HttpResponseMessage response = await client.PostAsync(authority, requestContent);

                if (response.IsSuccessStatusCode)
                {
                    string json = await response.Content.ReadAsStringAsync();
                    dynamic result = Newtonsoft.Json.JsonConvert.DeserializeObject(json);
                    return result.access_token;
                }
                else
                {
                    Console.WriteLine($"Failed to acquire access token. Status code: {response.StatusCode}");
                    string content = await response.Content.ReadAsStringAsync();
                    Console.WriteLine($"Response content: {content}");
                    return null;
                }
            }
        }
    }
}