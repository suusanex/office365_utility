using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Newtonsoft.Json;

namespace Office365UtilityLib.MicrosoftGraph
{
    public class Users
    {
        public class Value
        {
            [JsonProperty("@odata.type")]
            public string odatatype { get; set; }
            public string id { get; set; }
            public List<object> businessPhones { get; set; }
            public string displayName { get; set; }
            public string givenName { get; set; }
            public object jobTitle { get; set; }
            public string mail { get; set; }
            public string mobilePhone { get; set; }
            public object officeLocation { get; set; }
            public string preferredLanguage { get; set; }
            public string surname { get; set; }
            public string userPrincipalName { get; set; }
        }

        [JsonProperty("@odata.context")]
        public string odatacontext { get; set; }
        public List<Value> value { get; set; }
    }
}