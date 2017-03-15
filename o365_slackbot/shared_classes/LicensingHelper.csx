using System;
using System.IO;
using System.Net;
using Newtonsoft.Json.Linq;

public class LicensingHelper
{
    public static Array GetO365Skus(double apiVersion, string apiToken)
    {
        string skusList = null;
        var uri = $"https://graph.windows.net/myorganization/subscribedSkus?api-version={apiVersion}";
        Console.WriteLine(uri);

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

        return (Array)jObject["value"];

        //return skusList;
    }
}
