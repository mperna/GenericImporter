using Generic.Importer.Managers.Base;
using Generic.Importer.Interfaces;

namespace Generic.Importer.Managers
{
    public class Customer3PartManager : PartManagerBase, IPartManager
    {
        public Customer3PartManager(Database database) : base(database) { }
        
        #region Public Methods

        /// <summary>
        /// Retrieves the Customer3 part number ID for the given part number
        /// </summary>
        public int GetPartIdentifier(string partNumber)
        {
            return base.GetPartIdentifier("JSCustomer3ID", "Customer3PartNumber", partNumber);
        }

        /// <summary>
        /// Retrieves the Customer3 unit cost for the given part number.
        /// </summary>
        public decimal GetPartUnitCost(string partNumber)
        {
            return base.GetPartUnitCost("Customer3Price", "Customer3PartNumber", partNumber);
        }

        #endregion
    }
}
