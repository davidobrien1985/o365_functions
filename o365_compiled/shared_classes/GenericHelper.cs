using System;
using System.Net.Http;
using System.Text;
using Newtonsoft.Json;

namespace o365_compiled.shared_classes
{
    public class GenericHelper
    {
        public static string GetEnvironmentVariable(string name)
        {
            return System.Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.Process);
        }

        public static HttpResponseMessage SendMessageToSlack(string responseUri, object message)
        {
            var serializedPayload = JsonConvert.SerializeObject(message);
            HttpClient client = new HttpClient();
            var response = client.PostAsync(responseUri, new StringContent(serializedPayload, Encoding.UTF8, "application/json")).Result;
            return response;
        }
    }
}