using System;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs.Host;
using Newtonsoft.Json.Linq;
using o365_compiled.shared_classes;

namespace o365_compiled
{
    public class run
    {

        public static async Task<object> Run(HttpRequestMessage req, TraceWriter log)
        {
            log.Info($"C# HTTP trigger function processed a request. RequestUri={req.RequestUri}");

            string jsonContent = await req.Content.ReadAsStringAsync();
            log.Info(jsonContent);
            // assign the Slack payload "text" to be the UPN of the user that needs the license
            string username = (jsonContent.Split('&')[8]).Split('=')[1];
            log.Info(username);
            
            JArray skus = null;
            string skuId = null;
            string e3SkuId = null;
            string e1SkuId = null;
            double graphApiVersion = double.Parse(GetEnvironmentVariable("graphApiVersion"));
            string clientId = GetEnvironmentVariable("clientId");
            string clientSecret = GetEnvironmentVariable("clientSecret");
            string tenantId = GetEnvironmentVariable("tenantId");
            string res = null;
            JObject e3SkuObject = new JObject();
            JObject e1SkuObject = new JObject();

            // acquire Bearer Token for AD Application user through Graph API
            string token = AuthenticationHelperRest.AcquireTokenBySpn(tenantId, clientId, clientSecret);
            string bearerToken = "Bearer " + token;

            log.Info("Getting License SKUs...");
            // get information about all the O365 SKUs available
            skus = LicensingHelper.GetO365Skus(graphApiVersion, bearerToken);
            e1SkuObject = SubscriptionHelper.GetSkuId(skus, "STANDARDPACK");
            e3SkuObject = SubscriptionHelper.GetSkuId(skus, "ENTERPRISEPACK");

            e1SkuId = (string) e1SkuObject["skuId"];
            e3SkuId = (string) e3SkuObject["skuId"];

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
                    $"Successfully assigned license to {returnedUserName}";
            }
            else
            {
                // not enough licenses, notify user
                log.Info(
                    $"There are {purchasedLicenses} available E3 licenses and {usedLicenses} already used. No licenses available for E3. Please log on to portal.office.com and buy new licenses.");
                res =
                    $"There are {purchasedLicenses} available E3 licenses and {usedLicenses} already used. No licenses available for E3. Please log on to portal.office.com and buy new licenses.";
            }

            return res;
        }

        public static string GetEnvironmentVariable(string name)
        {
            return System.Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.Process);
        }
    }
}