using System;
using System.IO;
using System.Net;
using Newtonsoft.Json.Linq;

public class LicensingHelper
{
    public static JArray GetO365Skus(double apiVersion, string apiToken)
    {
        var uri = $"https://graph.windows.net/myorganization/subscribedSkus?api-version={apiVersion}";

        WebRequest request = WebRequest.Create(uri);
        request.Method = "GET";
        request.Headers.Add("Authorization", apiToken);

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

        JObject jObject = JObject.Parse(responseContent);

        return jObject.GetValue("value").Value<JArray>();
    }

    public static string SetO365LicensingInfo(double apiVersion, string apiToken, string userEmail, string addSkuId, string removeSkuId)
    {

        var uri = $"https://graph.windows.net/myorganization/users/{userEmail}/assignLicense?api-version={apiVersion}";

        var jsonPayload = $"{{\"addLicenses\": [{{\"disabledPlans\": [],\"skuId\": \"{addSkuId}\"}}],\"removeLicenses\": [\"{removeSkuId}\"]}}";

        string result = "";
        using (var client = new WebClient())
        {
            client.Headers[HttpRequestHeader.ContentType] = "application/json";
            client.Headers[HttpRequestHeader.Authorization] = apiToken;
            result = client.UploadString(uri, "POST", jsonPayload);
        }

        JObject resultJson = JObject.Parse(result);
        string username = resultJson.GetValue("userPrincipalName").Value<string>();

        return username;
    }
}
