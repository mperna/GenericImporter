
namespace Generic.Importer.Entities
{
	public class Customer2OrderLineItem
	{
		public double LineNumber { get; set; }

		public string ItemNumber { get; set; }
		public string ItemDescription { get; set; }
		public decimal ItemCost { get; set; }
		public int OrderQuantity { get; set; }
		public string NeedByDate { get; set; }

		public decimal Customer2ItemCost { get; set; }
		public int JSCustomer2PartNumberID { get; set; }
	}
}
