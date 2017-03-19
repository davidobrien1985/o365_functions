#r "Newtonsoft.Json"
#load "shared_classes/AuthenticationHelperRest.csx"
#load "shared_classes/LicensingHelper.csx"
using System.Net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;


public static Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
{
    log.Info($"C# HTTP trigger function processed a request. RequestUri={req.RequestUri}");

    var queryParams = req.GetQueryNameValuePairs()
        .ToDictionary(p => p.Key, p => p.Value, StringComparer.OrdinalIgnoreCase);

    HttpResponseMessage res = null;
    string name;
    if (queryParams.TryGetValue("name", out name))
    {
        res = new HttpResponseMessage(HttpStatusCode.OK)
        {
            Content = new StringContent("Hello " + name)
        };
    }
    else
    {
        res = new HttpResponseMessage(HttpStatusCode.BadRequest)
        {
            Content = new StringContent("Please pass a name on the query string")
        };
    }

    log.Info(name);

    double apiVersion = 1.6;
    JArray skus = null; 
    string skuId = null;
    string addSkuId = null;
    string removeSkuId = null;
    string clientId = GetEnvironmentVariable("clientId");
    string clientSecret = GetEnvironmentVariable("clientSecret");
    string tenantId = GetEnvironmentVariable("tenantId");

    string token = AuthenticationHelperRest.AcquireTokenBySpn(tenantId, clientId, clientSecret);
    string bearerToken = "Bearer " + token;
    
    log.Info("Getting License SKUs...");

    var uri = $"https://graph.windows.net/myorganization/subscribedSkus?api-version={apiVersion}";

    WebRequest request = WebRequest.Create(uri);
    request.Method = "GET";
    request.Headers.Add("Authorization", bearerToken);

    string responseContent = null;

    using (WebResponse response = request.GetResponse())
    {
        using (Stream stream = response.GetResponseStream())
        {
            if (stream != null)
                using (StreamReader sr99 = new StreamReader(stream))
                {
                    responseContent = sr99.ReadToEnd();
                }
        }
    }

    log.Info(responseContent);

    JObject jObject = JObject.Parse(responseContent);
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
                log.Info("There are {0} available E3 licenses and {1} already used.", purchasedLicenses, usedLicenses);
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

    LicensingHelper.SetO365LicensingInfo(apiVersion, bearerToken, name, addSkuId, removeSkuId);

    return Task.FromResult(res);
}

public static string GetEnvironmentVariable(string name)
{
return System.Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.Process);
}