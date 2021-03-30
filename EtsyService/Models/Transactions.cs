using System;
using System.Collections.Generic;
using System.Text;

namespace EtsyService.Models
{
    public class Transactions
    {
        public int? count { get; set; }
        public List<TransactionDetails> results { get; set; } = new List<TransactionDetails>();
    }
    public class TransactionDetails
    {
        public long? transaction_id { get; set; }
        public string title { get; set; }
        public string description { get; set; }
        public long? seller_user_id { get; set; }
        public long? buyer_user_id { get; set; }
        public long? creation_tsz { get; set; }
        public long? paid_tsz { get; set; }
        public long? shipped_tsz { get; set; }
        public string price { get; set; }
        public string currency_code { get; set; }
        public long? quantity { get; set; }
        public List<string> sku { get; set; } = new List<string>();
        public Listing Listing { get; set; }
    }
}
