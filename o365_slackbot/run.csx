#r "Newtonsoft.Json"
#load "shared_classes/AuthenticationHelperRest.csx"
#load "shared_classes/LicensingHelper.csx"
using System.Net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;


public static async Task<object> Run(HttpRequestMessage req, TraceWriter log)
{
    log.Info($"C# HTTP trigger function processed a request. RequestUri={req.RequestUri}");

    string jsonContent = await request.Content.ReadAsStringAsync();
    log.Info(jsonContent);
    string username = (jsonContent.Split('&')[8]).Split('=')[1];

    if ( username == null)
    {
        return req.CreateResponse(HttpStatusCode.BadRequest, new
        {
            error = "Please pass username in the input object"
        });
    }

    double apiVersion = 1.6;
    JArray skus = null; 
    string skuId = null;
    string e3SkuId = null;
    string e1SkuId = null;
    string clientId = GetEnvironmentVariable("clientId");
    string clientSecret = GetEnvironmentVariable("clientSecret");
    string tenantId = GetEnvironmentVariable("tenantId");

    string token = AuthenticationHelperRest.AcquireTokenBySpn(tenantId, clientId, clientSecret);
    string bearerToken = "Bearer " + token;
    
    log.Info("Getting License SKUs...");

    skus = LicensingHelper.GetO365Skus(apiVersion, bearerToken);

    for (int i = 0; i < skus.Count; i++)
    {

        JObject skuObject = (JObject)skus[i];

        skuId = (string)skuObject["skuId"];

        int usedLicenses = skuObject.GetValue("consumedUnits").Value<int>();
        int purchasedLicenses = skuObject.SelectToken(@"prepaidUnits.enabled").Value<int>();

        if ((string)skuObject["skuPartNumber"] == "ENTERPRISEPACK")
        {

            if (usedLicenses <= purchasedLicenses)
            {
                log.Info($"There are {purchasedLicenses} available E3 licenses and {usedLicenses} already used.");
                e3SkuId = skuId;
            }
            else
            {
                log.Info("No licenses available for E3. Please log on to portal.office.com and buy new licenses.");
            }
        }
        if ((string)skuObject["skuPartNumber"] == "STANDARDPACK")
        {
            e1SkuId = skuId;
        }
    }

    log.Info("Setting License...");
    LicensingHelper.SetO365LicensingInfo(apiVersion, bearerToken, username, e3SkuId, e1SkuId);

    var res = "";
    return res;
}

public static string GetEnvironmentVariable(string name)
{
return System.Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.Process);
}