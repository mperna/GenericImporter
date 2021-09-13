using System.Collections.Generic;

namespace Generic.Importer.Entities
{
    public class Customer2Order
    {
        public Customer2Order()
        {
            LineItems = new List<Customer2OrderLineItem>();
        }

        public string PurchaseOrderNumber { get; set; }
        public string PurchaseOrderDate { get; set;}
        public int Revision { get; set; }
        public decimal TotalCost { get; set; }
        public decimal TotalCustomer2Cost { get; set; }
        public string SalesOrderNumber { get; set; }
        
        public List<Customer2OrderLineItem> LineItems { get; set; }
    }
}
