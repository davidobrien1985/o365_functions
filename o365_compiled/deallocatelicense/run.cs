using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs.Host;

namespace deallocatelicense.deallocatelicense
{
    public class run
    {

        public static async Task<object> Run(HttpRequestMessage req, TraceWriter log)
        {

            log.Info("C# HTTP trigger function processed a request.");

            string res = null;

            return res;
        }
    }
}