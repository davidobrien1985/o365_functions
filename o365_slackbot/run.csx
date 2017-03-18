#load "shared_classes/AuthenticationHelperRest.csx"
#load "shared_classes/LicensingHelper.csx"
using System.Net;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
{
    log.Info($"C# HTTP trigger function processed a request. RequestUri={req.RequestUri}");

    // parse query parameter
    string user = req.user;

    // Get request body
    dynamic data = await req.Content.ReadAsAsync<object>();

    // Set name to query string or body data
    user = user ?? data?.user;

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
    LicensingHelper.SetO365LicensingInfo(apiVersion, bearerToken, user, addSkuId, removeSkuId);


return user == null
    ? req.CreateResponse(HttpStatusCode.BadRequest, "Please pass a name on the query string or in the request body")
    : req.CreateResponse(HttpStatusCode.OK, skus);
}

public static string GetEnvironmentVariable(string name)
{
return System.Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.Process);
}