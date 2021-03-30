using System;
using System.Collections.Generic;
using System.Text;

namespace EtsyService.Models
{
    public class UpdateInventory
    {
        public string sku { get; set; }
        public long product_id { get; set; }
        public List<object> property_values { get; set; } = new List<object>();
        public List<UpdateOffering> offerings { get; set; } = new List<UpdateOffering>();
    }

    public class UpdateOffering
    {
        public long offering_id { get; set; }
        public string price { get; set; }
        public int quantity { get; set; }
    }

}
