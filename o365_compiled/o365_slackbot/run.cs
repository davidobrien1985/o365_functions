﻿using System;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs.Host;
using Newtonsoft.Json.Linq;
using o365_compiled.o365_slackbot.shared_classes;

namespace o365_compiled.o365_slackbot
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
            
            double apiVersion = 1.6;
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

            skus = LicensingHelper.GetO365Skus(apiVersion, bearerToken);

            for (int i = 0; i < skus.Count; i++)
            {
                JObject skuObject = (JObject) skus[i];
                skuId = (string) skuObject["skuId"];

                if ((string) skuObject["skuPartNumber"] == "ENTERPRISEPACK")
                {
                    e3SkuObject = (JObject) skus[i];
                    e3SkuId = skuId;
                }
                if ((string) skuObject["skuPartNumber"] == "STANDARDPACK")
                {
                    e1SkuId = skuId;
                }
            }

            int usedLicenses = e3SkuObject.GetValue("consumedUnits").Value<int>();
            int purchasedLicenses = e3SkuObject.SelectToken(@"prepaidUnits.enabled").Value<int>();

            if (usedLicenses < purchasedLicenses)
            {
                log.Info("Setting License...");
                string returnedUserName =
                    LicensingHelper.SetO365LicensingInfo(apiVersion, bearerToken, username, e3SkuId, e1SkuId);
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