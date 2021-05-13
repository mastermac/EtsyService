using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EtsyService.Models
{
    public class OAuthProperties
    {
        public string oauth_consumer_key { get; set; }
        public string oauth_token { get; set; }
        public string oauth_signature_method { get; set; }
        public string oauth_timestamp { get; set; }
        public string oauth_nonce { get; set; }
        public string oauth_version { get; set; }
        public string oauth_signature { get; set; }
    }
}
