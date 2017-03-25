using System;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs.Host;
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
        public static string Run(payload req, TraceWriter log)
        {
            log.Info($"C# HTTP trigger function processed a request. Command used={req.Command}");

            double graphApiVersion = double.Parse(GetEnvironmentVariable("graphApiVersion"));
            string clientId = GetEnvironmentVariable("clientId");
            string clientSecret = GetEnvironmentVariable("clientSecret");
            string tenantId = GetEnvironmentVariable("tenantId");
            string allowedChannelName = GetEnvironmentVariable("allowedChannelName");
            string res = null;
            JObject e3SkuObject = new JObject();
            JObject e1SkuObject = new JObject();

            // assign the Slack payload "text" to be the UPN of the user that needs the license
            string username = req.Text;
            log.Info(username);

            // assign the Slack payload "channel_name" to the allowed channel name for this code to be called from
            string channelName = req.Channel_Name;
            if (channelName == allowedChannelName)
            {
                // acquire Bearer Token for AD Application user through Graph API
                string token = AuthenticationHelperRest.AcquireTokenBySpn(tenantId, clientId, clientSecret);
                string bearerToken = "Bearer " + token;

                log.Info("Getting License SKUs...");
                // get information about all the O365 SKUs available
                JArray skus = LicensingHelper.GetO365Skus(graphApiVersion, bearerToken);
                e1SkuObject = SubscriptionHelper.FilterSkusByPartNumber(skus, "STANDARDPACK");
                e3SkuObject = SubscriptionHelper.FilterSkusByPartNumber(skus, "ENTERPRISEPACK");

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
                    res =
                        $"There are {purchasedLicenses} available E3 licenses and {usedLicenses} already used. You have just used one more." +
                        $"Successfully assigned *E3* license to {returnedUserName}";
                }
                else
                {
                    // not enough licenses, notify user
                    log.Info(
                        $"There are {purchasedLicenses} available E3 licenses and {usedLicenses} already used. No licenses available for E3. Please log on to portal.office.com and buy new licenses.");
                    res =
                        $"There are {purchasedLicenses} available E3 licenses and {usedLicenses} already used. No licenses available for E3. Please log on to portal.office.com and buy new licenses.";
                }
            }
            else
            {
                res =
                    $"The channel {channelName} is not allowed to execute this function. Sorry. Get in touch with your O365 admin.";
            }
            return res;
        }

        public static string GetEnvironmentVariable(string name)
        {
            return System.Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.Process);
        }
    }

    public class Deallocatelicense
    {
        public static string Run(payload req, TraceWriter log)
        {

            log.Info($"C# HTTP trigger function processed a request.Requested by= {req.User_Name}");
            double graphApiVersion = double.Parse(GetEnvironmentVariable("graphApiVersion"));
            string clientId = GetEnvironmentVariable("clientId");
            string clientSecret = GetEnvironmentVariable("clientSecret");
            string tenantId = GetEnvironmentVariable("tenantId");
            string allowedChannelName = GetEnvironmentVariable("allowedChannelName");
            JObject e3SkuObject = new JObject();
            JObject e1SkuObject = new JObject();
            string res = null;

            string username = req.Text;

            // assign the Slack payload "channel_name" to the allowed channel name for this code to be called from
            string channelName = req.Channel_Name;
            if (channelName == allowedChannelName)
            {
                // acquire Bearer Token for AD Application user through Graph API
                string token = AuthenticationHelperRest.AcquireTokenBySpn(tenantId, clientId, clientSecret);
                string bearerToken = "Bearer " + token;

                log.Info("Getting License SKUs...");
                // get information about all the O365 SKUs available
                JArray skus = LicensingHelper.GetO365Skus(graphApiVersion, bearerToken);
                e1SkuObject = SubscriptionHelper.FilterSkusByPartNumber(skus, "STANDARDPACK");
                e3SkuObject = SubscriptionHelper.FilterSkusByPartNumber(skus, "ENTERPRISEPACK");

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
                        LicensingHelper.SetO365LicensingInfo(graphApiVersion, bearerToken, username, e1SkuId, e3SkuId);
                    res =
                        $"There are {purchasedLicenses} available E1 licenses and {usedLicenses} already used. You have just used one more." +
                        $"Successfully assigned *E1* license to {returnedUserName}";
                }
                else
                {
                    // not enough licenses, notify user
                    log.Info(
                        $"There are {purchasedLicenses} available E1 licenses and {usedLicenses} already used. No licenses available for E1. Please log on to portal.office.com and buy new licenses.");
                    res =
                        $"There are {purchasedLicenses} available E1 licenses and {usedLicenses} already used. No licenses available for E1. Please log on to portal.office.com and buy new licenses.";
                }
            }

            return res;
        }

        public static string GetEnvironmentVariable(string name)
        {
            return System.Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.Process);
        }
    }

    public class GetLicenseInfo
    {
        public static string Run(payload req, TraceWriter log)
        {
            log.Info($"C# HTTP trigger function processed a request.Command used= {req.Command}");
            double graphApiVersion = double.Parse(GetEnvironmentVariable("graphApiVersion"));
            string clientId = GetEnvironmentVariable("clientId");
            string clientSecret = GetEnvironmentVariable("clientSecret");
            string tenantId = GetEnvironmentVariable("tenantId");
            string allowedChannelName = GetEnvironmentVariable("allowedChannelName");
            string res = null;

            // assign the Slack payload "text" to be the UPN of the user that needs the license
            string username = req.Text;

            log.Info(username);
            string encUserName = Uri.UnescapeDataString(username);

            // assign the Slack payload "channel_name" to the allowed channel name for this code to be called from
            string channelName = req.Channel_Name;
            if (channelName == allowedChannelName)
            {
                // acquire Bearer Token for AD Application user through Graph API
                string token = AuthenticationHelperRest.AcquireTokenBySpn(tenantId, clientId, clientSecret);
                string bearerToken = "Bearer " + token;
                string skuPartNumber =
                    LicensingHelper.GetUserLicenseInfo(graphApiVersion, encUserName, bearerToken, log);

                res = $"{encUserName} is licensed with the {skuPartNumber} license.";
            }
            return res;

        }

        public static string GetEnvironmentVariable(string name)
        {
            return System.Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.Process);
        }
    }
}