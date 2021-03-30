using System;
using System.Collections.Generic;
using System.Text;

namespace EtsyService.Models
{
    public class Listing
    {
        public int listing_id { get; set; }
        public string state { get; set; }
        public string title { get; set; }
        public string description { get; set; }
        public string price { get; set; }
        public string currency_code { get; set; }
        public int quantity { get; set; }
        public string url { get; set; }
        public List<string> sku { get; set; } = new List<string>();
    }
}
