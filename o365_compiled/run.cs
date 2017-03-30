using System;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using o365_compiled.shared_classes;

namespace o365_compiled
{
    public class payload
    {
        public string Token { get; set; }
        public string Team_Id { get; set; }
        public string Team_Domain { get; set; }
        public string Channel_Id { get; set; }
        public string Channel_Name { get; set; }
        public string User_Id { get; set; }
        public string User_Name { get; set; }
        public string Command { get; set; }
        public string Text { get; set; }
        public string Response_Url { get; set; }

    }

    public class run
    {
        public static async Task<string> Run(HttpRequestMessage req, TraceWriter log, IAsyncCollector<payload> allocatee3o365, IAsyncCollector<payload> geto365userlicense, IAsyncCollector<payload> deallocatee3o365, ICollector<string> outputDocument)
        {
            log.Info($"C# HTTP trigger function processed a request. Command used={req.RequestUri}");

            string jsonContent = await req.Content.ReadAsStringAsync();
            log.Info(jsonContent);
            string allowedChannelName = GenericHelper.GetEnvironmentVariable("allowedChannelName");
            string res = null;
            
            // assign the Slack payload "channel_name" to the allowed channel name for this code to be called from
            string channelName = (jsonContent.Split('&')[4]).Split('=')[1];
            if (channelName == allowedChannelName)
            {
                
                payload json = new payload
                {
                    Token = (jsonContent.Split('&')[0]).Split('=')[1],
                    Team_Id = (jsonContent.Split('&')[1]).Split('=')[1],
                    Team_Domain = (jsonContent.Split('&')[2]).Split('=')[1],
                    Channel_Id = (jsonContent.Split('&')[3]).Split('=')[1],
                    Channel_Name = (jsonContent.Split('&')[4]).Split('=')[1],
                    User_Id = (jsonContent.Split('&')[5]).Split('=')[1],
                    User_Name = (jsonContent.Split('&')[6]).Split('=')[1],
                    Command = (jsonContent.Split('&')[7]).Split('=')[1],
                    Text = (jsonContent.Split('&')[8]).Split('=')[1],
                    Response_Url = (jsonContent.Split('&')[9]).Split('=')[1]
                };

                string document = JsonConvert.SerializeObject(json);
                outputDocument.Add(document);

                string command = Uri.EscapeDataString(json.Command);

                if (command == "%252Fallocatee3o365")
                {
                    // add to allocateE3O365 queue
                    await allocatee3o365.AddAsync(json);
                }
                else if (command == "%252Fdeallocatee3o365")
                {
                    // add to deallocatee3o365 queue
                    await deallocatee3o365.AddAsync(json);
                }
                else if (command == "%252Fgeto365userlicense")
                {
                    // add to geto365userlicense queue
                    await geto365userlicense.AddAsync(json);
                }
                
                res = $"Hey, {json.User_Name}, I'm working on assigning the license. I'll let you know when I'm done...";
            }
            return res;
        }
    }

    
    public class AllocateE3License
    {
        public static void Run(payload allocatee3o365, TraceWriter log)
        {
            log.Info($"C# Queue trigger function processed: {allocatee3o365.User_Name}");
            double graphApiVersion = double.Parse(GenericHelper.GetEnvironmentVariable("graphApiVersion"));
            string clientId = GenericHelper.GetEnvironmentVariable("clientId");
            string clientSecret = GenericHelper.GetEnvironmentVariable("clientSecret");
            string tenantId = GenericHelper.GetEnvironmentVariable("tenantId");

            // assign the Slack payload "text" to be the UPN of the user that needs the license
            string username = allocatee3o365.Text;
            log.Info(username);

            // acquire Bearer Token for AD Application user through Graph API
            string token = AuthenticationHelperRest.AcquireTokenBySpn(tenantId, clientId, clientSecret);
            string bearerToken = "Bearer " + token;

            log.Info("Getting License SKUs...");
            // get information about all the O365 SKUs available
            JArray skus = LicensingHelper.GetO365Skus(graphApiVersion, bearerToken);
            JObject e1SkuObject = SubscriptionHelper.FilterSkusByPartNumber(skus, "STANDARDPACK");
            JObject e3SkuObject = SubscriptionHelper.FilterSkusByPartNumber(skus, "ENTERPRISEPACK");

            string e1SkuId = (string) e1SkuObject["skuId"];
            string e3SkuId = (string) e3SkuObject["skuId"];

            int usedLicenses = e3SkuObject.GetValue("consumedUnits").Value<int>();
            int purchasedLicenses = e3SkuObject.SelectToken(@"prepaidUnits.enabled").Value<int>();

            // check if enough licenses available
            if (usedLicenses < purchasedLicenses)
            {
                // enough licenses, so do it
                log.Info("Setting License...");
                string returnedUserName =
                    LicensingHelper.SetO365LicensingInfo(graphApiVersion, bearerToken, username, e3SkuId, e1SkuId);

                var uri = Uri.UnescapeDataString(allocatee3o365.Response_Url);

                var jsonPayload = new
                {
                    text =
                    $"There are {purchasedLicenses} available E3 licenses and {usedLicenses} already used. You have just used one more. Successfully assigned *E3* license to {returnedUserName}"
                };

                GenericHelper.SendMessageToSlack(uri, jsonPayload);
            }
            else
            {
                // not enough licenses, notify user
                log.Info(
                    $"There are {purchasedLicenses} available E3 licenses and {usedLicenses} already used. No licenses available for E3. Please log on to portal.office.com and buy new licenses.");
                var uri = Uri.UnescapeDataString(allocatee3o365.Response_Url);

                log.Info(uri);
                var jsonPayload = new
                {
                    text =
                    $"There are {purchasedLicenses} available E3 licenses and {usedLicenses} already used. No licenses available for E3. Please log on to portal.office.com and buy new licenses."
                };

                GenericHelper.SendMessageToSlack(uri, jsonPayload);
            }
        }
    }

    public class Deallocatelicense
    {
        public static void Run(payload deallocatee3o365, TraceWriter log)
        {

            log.Info($"C# HTTP trigger function processed a request. Command used={deallocatee3o365.Command}");
            double graphApiVersion = double.Parse(GenericHelper.GetEnvironmentVariable("graphApiVersion"));
            string clientId = GenericHelper.GetEnvironmentVariable("clientId");
            string clientSecret = GenericHelper.GetEnvironmentVariable("clientSecret");
            string tenantId = GenericHelper.GetEnvironmentVariable("tenantId");
            string allowedChannelName = GenericHelper.GetEnvironmentVariable("allowedChannelName");

            // assign the Slack payload "text" to be the UPN of the user that needs the license
            string username = deallocatee3o365.Text;
            log.Info(username);

            // assign the Slack payload "channel_name" to the allowed channel name for this code to be called from
            string channelName = deallocatee3o365.Channel_Name;
            if (channelName == allowedChannelName && deallocatee3o365.User_Name == "david")
            {
                // acquire Bearer Token for AD Application user through Graph API
                string token = AuthenticationHelperRest.AcquireTokenBySpn(tenantId, clientId, clientSecret);
                string bearerToken = "Bearer " + token;

                log.Info("Getting License SKUs...");
                // get information about all the O365 SKUs available
                JArray skus = LicensingHelper.GetO365Skus(graphApiVersion, bearerToken);
                JObject e1SkuObject = SubscriptionHelper.FilterSkusByPartNumber(skus, "STANDARDPACK");
                JObject e3SkuObject = SubscriptionHelper.FilterSkusByPartNumber(skus, "ENTERPRISEPACK");
                string e1SkuId = (string) e1SkuObject["skuId"];
                string e3SkuId = (string) e3SkuObject["skuId"];

                int usedLicenses = e1SkuObject.GetValue("consumedUnits").Value<int>();
                int purchasedLicenses = e1SkuObject.SelectToken(@"prepaidUnits.enabled").Value<int>();

                // check if enough licenses available
                if (usedLicenses < purchasedLicenses)
                {
                    // enough licenses, so do it
                    log.Info("Setting License...");
                    string returnedUserName =
                        LicensingHelper.SetO365LicensingInfo(graphApiVersion, bearerToken, username, e1SkuId,
                            e3SkuId);

                    var uri = Uri.UnescapeDataString(deallocatee3o365.Response_Url);
                    var jsonPayload = new
                    {
                        text =
                        $"There are {purchasedLicenses} available E1 licenses and {usedLicenses} already used. You have just used one more." +
                        $"Successfully assigned *E1* license to {returnedUserName}"
                    };

                    GenericHelper.SendMessageToSlack(uri, jsonPayload);
                }
                else
                {
                    // not enough licenses, notify user
                    log.Info(
                        $"There are {purchasedLicenses} available E1 licenses and {usedLicenses} already used. No licenses available for E1. Please log on to portal.office.com and buy new licenses.");
                    
                    var uri = Uri.UnescapeDataString(deallocatee3o365.Response_Url);
                    var jsonPayload = new
                    {
                        text =
                        $"There are {purchasedLicenses} available E1 licenses and {usedLicenses} already used. No licenses available for E1. Please log on to portal.office.com and buy new licenses."
                    };

                    GenericHelper.SendMessageToSlack(uri, jsonPayload);
                }
            }
        }
    }

    public class GetLicenseInfo
    {
        public static void Run(payload geto365userlicense, TraceWriter log)
        {
            log.Info($"C# HTTP trigger function processed a request. Command used={geto365userlicense.Command}");
            double graphApiVersion = double.Parse(GenericHelper.GetEnvironmentVariable("graphApiVersion"));
            string clientId = GenericHelper.GetEnvironmentVariable("clientId");
            string clientSecret = GenericHelper.GetEnvironmentVariable("clientSecret");
            string tenantId = GenericHelper.GetEnvironmentVariable("tenantId");
            string allowedChannelName = GenericHelper.GetEnvironmentVariable("allowedChannelName");

            // assign the Slack payload "text" to be the UPN of the user that needs the license
            string username = geto365userlicense.Text;
            string encUserName = Uri.UnescapeDataString(username);

            log.Info(encUserName);
            // assign the Slack payload "channel_name" to the allowed channel name for this code to be called from
            string channelName = geto365userlicense.Channel_Name;
            if (channelName == allowedChannelName)
            {
                // acquire Bearer Token for AD Application user through Graph API
                string token = AuthenticationHelperRest.AcquireTokenBySpn(tenantId, clientId, clientSecret);
                string bearerToken = "Bearer " + token;
                string skuPartNumber =
                    LicensingHelper.GetUserLicenseInfo(graphApiVersion, encUserName, bearerToken, log);

                var uri = Uri.UnescapeDataString(geto365userlicense.Response_Url);
                var jsonPayload = new
                {
                    text = $"{encUserName} is licensed with the {skuPartNumber} license."
                };

                GenericHelper.SendMessageToSlack(uri, jsonPayload);
            }
            
        }
    }
}