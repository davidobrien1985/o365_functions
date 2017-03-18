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

        Console.Write(responseContent);

        JObject jObject = JObject.Parse(responseContent);

        Console.Write(jObject["value"]);
        return (JArray)jObject["value"];
    }

    public static string SetO365LicensingInfo(double apiVersion, string apiToken, string userEmail, string addSkuId, string removeSkuId)
    {

        var uri = $"https://graph.windows.net/myorganization/users/{userEmail}/assignLicense?api-version={apiVersion}";

        // 6fd2c87f-b296-42f0-b197-1e91e994b900
        // 18181a46-0d4e-45cd-891e-60aabd171b4e

        var jsonPayload = $"{{\"addLicenses\": [{{\"disabledPlans\": [],\"skuId\": \"{addSkuId}\"}}],\"removeLicenses\": [\"{removeSkuId}\"]}}";


        string result = "";
        using (var client = new WebClient())
        {
            client.Headers[HttpRequestHeader.ContentType] = "application/json";
            client.Headers[HttpRequestHeader.Authorization] = apiToken;
            result = client.UploadString(uri, "POST", jsonPayload);
        }
        Console.WriteLine(result);

        return result;
    }
}
