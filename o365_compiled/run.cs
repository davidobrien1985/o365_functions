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
            string username = (jsonContent.Split('&')[8]).Split('=')[1];
            log.Info(username);
            
            double graphApiVersion = double.Parse(GetEnvironmentVariable("graphApiVersion"));
            JArray skus = null;
            string skuId = null;
            string e3SkuId = null;
            string e1SkuId = null;
            string clientId = GetEnvironmentVariable("clientId");
            string clientSecret = GetEnvironmentVariable("clientSecret");
            string tenantId = GetEnvironmentVariable("tenantId");
            string res = null;
            JObject e3SkuObject = new JObject();

            string token = AuthenticationHelperRest.AcquireTokenBySpn(tenantId, clientId, clientSecret);
            string bearerToken = "Bearer " + token;

            log.Info("Getting License SKUs...");

            skus = LicensingHelper.GetO365Skus(graphApiVersion, bearerToken);
            e1SkuId = SubscriptionHelper.GetSkuId(skus, "STANDARDPACK");
            e3SkuId = SubscriptionHelper.GetSkuId(skus, "ENTERPRISEPACK");

            if ((e1SkuId  == "00000000000000000000000000") || (e3SkuId == "00000000000000000000000000"))
            {
                throw new SystemException("Unable to find skuIDs");
            }


            int usedLicenses = e3SkuObject.GetValue("consumedUnits").Value<int>();
            int purchasedLicenses = e3SkuObject.SelectToken(@"prepaidUnits.enabled").Value<int>();

            if (usedLicenses < purchasedLicenses)
            {
                log.Info("Setting License...");
                string returnedUserName =
                    LicensingHelper.SetO365LicensingInfo(graphApiVersion, bearerToken, username, e3SkuId, e1SkuId);
                res =
                    $"There are {purchasedLicenses} available E3 licenses and {usedLicenses} already used. You have just used one more." +
                    $"Successfully assigned license to {returnedUserName}";
            }
            else
            {
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