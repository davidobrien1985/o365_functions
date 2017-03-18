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

    log.Info($"{queryParams}");

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

    skus = LicensingHelper.GetO365Skus(apiVersion, bearerToken);

    for (int i = 0; i < skus.Count; i++)
    {

        JObject skuObject = (JObject)skus[i];

        skuId = (string)skuObject["skuId"];

        if ((string)skuObject["skuPartNumber"] == "ENTERPRISEPACK")
        {
            if ((int)skuObject["consumedUnits"] <= (int)skuObject["prepaidUnits.enabled"])
            {
                addSkuId = skuId;
            }
        }
        if ((string)skuObject["skuPartNumber"] == "STANDARDPACK")
        {
            removeSkuId = skuId;
        }
    }

    Console.WriteLine(addSkuId);
    Console.WriteLine(removeSkuId);
    LicensingHelper.SetO365LicensingInfo(apiVersion, bearerToken, name, addSkuId, removeSkuId);


    return Task.FromResult(res);
}

public static string GetEnvironmentVariable(string name)
{
return System.Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.Process);
}