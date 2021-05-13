using System;
using System.Collections.Generic;
using System.Text;

namespace EtsyService.Models
{
    public class GetInventory
    {
        public int count { get; set; } = 0;
        public Results results { get; set; }
        public long price_on_property { get; set; }
        public long quantity_on_property { get; set; }
        public long sku_on_property { get; set; }
    }
    public class Results
    {
        public List<Product> products { get; set; }
    }
    public class Offering
    {
        public long offering_id { get; set; }
        public object price { get; set; }
        public int quantity { get; set; }
        public int is_enabled { get; set; }
        public int is_deleted { get; set; }
    }
    public class Price
    {
        public int amount { get; set; }
        public int divisor { get; set; }
        public string currency_code { get; set; }
        public string currency_formatted_short { get; set; }
        public string currency_formatted_long { get; set; }
        public string currency_formatted_raw { get; set; }
    }

    public class Product
    {
        public long product_id { get; set; }
        public string sku { get; set; }
        public List<object> property_values { get; set; }
        public List<Offering> offerings { get; set; }
    }
}
