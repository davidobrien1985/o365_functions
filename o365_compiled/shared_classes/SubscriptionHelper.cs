using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Newtonsoft.Json.Linq;

namespace o365_compiled.shared_classes
{
    public class SubscriptionHelper
    {
        public static JObject GetSkuId (JArray skus, string skuPartNumber)
        {
            foreach (JToken sku in skus)
            {
                JObject skuObject = (JObject)sku;

                if ((string)skuObject["skuPartNumber"] == skuPartNumber)
                {
                    skuObject = (JObject)sku;
                    return skuObject;
                }
            }
            

            return new JObject();
        }
    }
}