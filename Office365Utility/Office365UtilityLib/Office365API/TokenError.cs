using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Office365UtilityLib.Office365API
{
    public class TokenError
    {
        public string error { get; set; }
        public string error_description { get; set; }
        public List<int> error_codes { get; set; }
        public string timestamp { get; set; }
        public string trace_id { get; set; }
        public string correlation_id { get; set; }

        public override string ToString()
        {
            return $"{nameof(TokenError)}:{nameof(error)}={error}, {nameof(error_description)}={error_description}, {nameof(timestamp)}={timestamp}, {nameof(trace_id)}={trace_id}, {nameof(correlation_id)}={correlation_id}, {nameof(error_codes)}={string.Join(",", error_codes)}";
        }
    }
}