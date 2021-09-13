using System.Collections.Generic;

namespace Generic.Importer.Entities
{
    public class Customer1Order
    {
        public Customer1Order()
        {
            LineItems = new List<Customer1OrderLineItem>();
        }

        public double OrderNumber { get; set; }
        public int Revision { get; set; }
        public double TotalCost { get; set; }
        public double TotalCustomer1Cost { get; set; }
        public string ImportDate { get; set; }

        public List<Customer1OrderLineItem> LineItems { get; set; }
    }
}
