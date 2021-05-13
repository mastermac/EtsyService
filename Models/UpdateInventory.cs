using System;
using System.Collections.Generic;
using System.Text;

namespace EtsyService.Models
{
    public class UpdateInventory
    {
        public string sku { get; set; }
        public long product_id { get; set; }
        public List<PropValues> property_values { get; set; } = new List<PropValues>();
        public List<UpdateOffering> offerings { get; set; } = new List<UpdateOffering>();
    }
    public class UpdateInventory1
    {
        public List<Product> Products { get; set; } = new List<Product>();
    }

    public class UpdateOffering
    {
        public long offering_id { get; set; }
        public string price { get; set; }
        public int quantity { get; set; }
    }
    public class PropValues
    {
        public long property_id { get; set; }
        public string property_name { get; set; }
        public long scale_id { get; set; }
        public string scale_name { get; set; }
        public List<string> values { get; set; } = new List<string>();
        public List<long> value_ids { get; set; } = new List<long>();

    }

}
