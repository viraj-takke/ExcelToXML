using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToXML.Models
{
    public class OrderData
    {
        public string UniqueIdentity { get; set; }
        public string OrderId { get; set; }
        public DateTime OrderDate { get; set; }
        public string OrderPerson { get; set; }
        public string ShipToName { get; set; }
        public string ShipToAddress { get; set; }
        public string ShipToCity { get; set; }
        public string ShipToRegion { get; set; }
        public List<ItemData> Items { get; set; } = new List<ItemData>();
    }

    public class ItemData
    {
        public string Title { get; set; }
        public string Note { get; set; }
        public int Quantity { get; set; }
        public decimal Price { get; set; }
        public decimal Total { get; set; }
    }

    public class ContactInfo
    {
        public string Name { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string Region { get; set; }
    }
}
