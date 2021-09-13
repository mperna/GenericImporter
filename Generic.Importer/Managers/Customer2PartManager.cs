using Generic.Importer.Managers.Base;
using Generic.Importer.Interfaces;

namespace Generic.Importer.Managers
{
    public class Customer2PartManager : PartManagerBase, IPartManager
    {
        public Customer2PartManager(Database database) : base(database) { }

        #region Public Methods

        /// <summary>
        /// Retrieves the Customer1 part number ID for the given part number
        /// </summary>
        public int GetPartIdentifier(string partNumber)
        {
            return base.GetPartIdentifier("JSCustomer2ID", "Customer2PartNumber", partNumber);
        }

        /// <summary>
        /// Retrieves the Customer1 unit cost for the given part number.
        /// </summary>
        public decimal GetPartUnitCost(string partNumber)
        {
            return base.GetPartUnitCost("Customer2Price", "Customer2PartNumber", partNumber);
        }

        #endregion
     }
}
