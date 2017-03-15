using System;
using System.IO;
using System.Net;
using System.Text;
using Newtonsoft.Json.Linq;

public class AuthenticationHelperRest
{

    const string ArmResource = "https://graph.windows.net/";
    const string TokenEndpoint = "https://login.windows.net/{0}/oauth2/token";
    const string SpnPayload = "resource={0}&client_id={1}&grant_type=client_credentials&client_secret={2}";

    public static string AcquireTokenBySpn(string tenantId, string clientId, string clientSecret)
    {
        var payload = String.Format(SpnPayload, WebUtility.UrlEncode(ArmResource), WebUtility.UrlEncode(clientId),
            WebUtility.UrlEncode(clientSecret));

        byte[] data = Encoding.ASCII.GetBytes(payload);

        var address = String.Format(TokenEndpoint, tenantId);
        WebRequest request = WebRequest.Create(address);
        request.Method = "POST";
        request.ContentType = "application/x-www-form-urlencoded";
        request.ContentLength = data.Length;
        using (Stream stream = request.GetRequestStream())
        {
            stream.Write(data, 0, data.Length);
        }

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

        return (string)jObject["access_token"];
    }
}
