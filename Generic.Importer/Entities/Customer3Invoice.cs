using System.Collections.Generic;

namespace Generic.Importer.Entities
{
    public class Customer3Invoice
    {
        public Customer3Invoice()
        {
            LineItems = new List<Customer3InvoiceLineItem>();
        }

        public string SalesOrderNumber { get; set; }
        public string SalesOrderStatus { get; set; }
        public int Revision { get; set; }
        public int PurchaseOrder { get; set; }
        public decimal TotalCost { get; set; }
        public decimal TotalCustomer3Cost { get; set; }
        public string NeedByDate { get; set; }
        public string RouteCode { get; set; }

        public List<Customer3InvoiceLineItem> LineItems { get; set; }
    }
}
