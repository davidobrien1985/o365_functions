#r "Newtonsoft.Json"
#load "shared_classes/AuthenticationHelperRest.csx"
#load "shared_classes/LicensingHelper.csx"
using System.Net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;


public static async Task<object> Run(HttpRequestMessage req, TraceWriter log)
{
    log.Info($"C# HTTP trigger function processed a request. RequestUri={req.RequestUri}");

    string jsonContent = await req.Content.ReadAsStringAsync();
    dynamic data = JsonConvert.DeserializeObject(jsonContent);
    
    log.Info(jsonContent);

    if (data.channel == null || data.username == null || data.text == null || data.icon_url == null)
    {
        return req.CreateResponse(HttpStatusCode.BadRequest, new
        {
            error = "Please pass channel/username/text/icon_url properties in the input object"
        });
    }

    var payload = new
    {
        channel = data.channel,
        username = data.username,
        text = data.text,
        icon_url = data.icon_url,
    };

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
    // fix jsoncontent for actual email address !!!!!!!!!!!!!!!!!!!!!!!!!!!
    LicensingHelper.SetO365LicensingInfo(apiVersion, bearerToken, jsonContent, e3SkuId, e1SkuId);

    var jsonString = JsonConvert.SerializeObject(payload);
    //using (var client = new HttpClient())
    //{
    //    var res = await client.PostAsync(_slackWebhookUrl, new FormUrlEncodedContent(new[]
    //    {
    //        new KeyValuePair<string, string>("payload", jsonString)
    //    }));
    //    return req.CreateResponse(res.StatusCode, new
    //    {
    //        body = $"Send to Slack for following. text : {data.text}",
    //    });
    //}
}

public static string GetEnvironmentVariable(string name)
{
return System.Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.Process);
}