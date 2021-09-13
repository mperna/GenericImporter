using Generic.Importer.Managers.Base;
using Generic.Importer.Interfaces;

namespace Generic.Importer.Managers
{
    public class Customer1PartManager : PartManagerBase, IPartManager
    {
        public Customer1PartManager(Database database) : base(database) { }

        #region Public Methods

        /// <summary>
        /// Retrieves the Customer1 part number ID for the given part number
        /// </summary>
        public int GetPartIdentifier(string partNumber)
        {
            return base.GetPartIdentifier("JSCustomer1ID", "Customer1PartNumber", partNumber);
        }

        /// <summary>
        /// Retrieves the Customer1 unit cost for the given part number.
        /// </summary>
        public decimal GetPartUnitCost(string partNumber)
        {
            return base.GetPartUnitCost("Customer1Price", "Customer1PartNumber", partNumber);
        }

        #endregion
    }
}
