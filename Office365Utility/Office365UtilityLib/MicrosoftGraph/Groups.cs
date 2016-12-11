using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Newtonsoft.Json;

namespace Office365UtilityLib.MicrosoftGraph
{
    public class Groups
    {

        public class Value
        {
            public string id { get; set; }
            public object classification { get; set; }
            public string createdDateTime { get; set; }
            public string description { get; set; }
            public string displayName { get; set; }
            public List<string> groupTypes { get; set; }
            public string mail { get; set; }
            public bool mailEnabled { get; set; }
            public string mailNickname { get; set; }
            public object onPremisesLastSyncDateTime { get; set; }
            public object onPremisesSecurityIdentifier { get; set; }
            public object onPremisesSyncEnabled { get; set; }
            public List<object> proxyAddresses { get; set; }
            public string renewedDateTime { get; set; }
            public bool securityEnabled { get; set; }
            public string visibility { get; set; }
        }

        [JsonProperty("@odata.context")]
        public string odatacontext { get; set; }
        public List<Value> value { get; set; }
    }

}
