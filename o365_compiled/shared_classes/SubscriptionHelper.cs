using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Newtonsoft.Json.Linq;

namespace o365_compiled.shared_classes
{
    public class SubscriptionHelper
    {
        public static string GetSkuId (JArray skus, string skuPartNumber)
        {
            string skuId = null;

            foreach (JToken sku in skus)
            {
                JObject skuObject = (JObject)sku;
                skuId = (string)skuObject["skuId"];

                if ((string)skuObject["skuPartNumber"] == skuPartNumber)
                {
                    skuObject = (JObject)sku;
                    return skuId;
                }
            }
            if (skuId == null)
            {
                return "00000000000000000000000000";
            }

            return "00000000000000000000000000";
        }
    }
}