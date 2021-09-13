using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using Generic.Importer.Managers.Base;
using Generic.Importer.Entities;

namespace Generic.Importer.Managers
{
    public class Customer1OrderManager : OrderManagerBase
    {
        public Customer1OrderManager(Database database) : base(database) { }

        #region Public Methods

        /// <summary>
        /// Determines if the current order represents data already in the database
        /// </summary>
        public bool Exists(Customer1Order order)
        {
            return base.Exists("Customer1 Orders", "Customer1PONumber", order.OrderNumber.ToString());
        }

        public void Delete(Customer1Order order)
        {
            base.Delete("Customer1 Orders", "Customer1 Order Details", "Customer1PONumber", order.OrderNumber.ToString());
        }

        /// <summary>
        /// Purging all forecast records by removing all data from Customer1 Orders/Order Details which
        /// have a purchase order starting with 1.
        /// </summary>
        public void PurgeForecasts()
        {
            if ((_database.Connection != null) && (_database.Connection.State == ConnectionState.Open))
            {
                //Searching for an existing invoice in the database with this purchase order number.

                using (OleDbCommand command = _database.Connection.CreateCommand())
                {
                    command.CommandType = CommandType.Text;
                    command.CommandText = "DELETE FROM [Customer1 Order Details] WHERE Customer1PONumber LIKE '1%'";

                    command.ExecuteScalar();
                }

                //Delete the main purchase order record.

                using (OleDbCommand command = _database.Connection.CreateCommand())
                {

                    command.CommandType = CommandType.Text;
                    command.CommandText = "DELETE FROM [Customer1 Orders] WHERE Customer1PONumber LIKE '1%'";

                    command.ExecuteScalar();
                }
            }
        }

        public void Save(Customer1Order order)
        {
            if ((order != null) && (_database.Connection != null) && (_database.Connection.State == ConnectionState.Open))
            {
                bool recordExists = Exists(order);

                //If this order already exists in the database, delete it and its detail records before
                //attempting to re-insert it into the database with the current data.

                if (recordExists == true)
                {
                    Delete(order);
                }

                SaveCustomer1Order(order);
                SaveCustomer1OrderDetails(order);
            }
        }

        #endregion

        #region Private Methods

        private void SaveCustomer1Order(Customer1Order order)
        {
            if ((order != null) && (_database.Connection != null) && (_database.Connection.State == ConnectionState.Open))
            {
                using (OleDbCommand command = _database.Connection.CreateCommand())
                {
                    DbParameter parameter1 = command.CreateParameter();
                    parameter1.ParameterName = "TotalOrderDollars";
                    parameter1.DbType = DbType.Decimal;
                    parameter1.Value = order.TotalCost;
                    command.Parameters.Add(parameter1);

                    DbParameter parameter2 = command.CreateParameter();
                    parameter2.ParameterName = "Customer1PONumber";
                    parameter2.DbType = DbType.String;
                    parameter2.Value = order.OrderNumber;
                    command.Parameters.Add(parameter2);

                    DbParameter parameter3 = command.CreateParameter();
                    parameter3.ParameterName = "PODate";
                    parameter3.DbType = DbType.DateTime;
                    parameter3.Value = order.ImportDate;
                    command.Parameters.Add(parameter3);

                    command.CommandType = CommandType.Text;
                    command.CommandText = "INSERT INTO [Customer1 Orders] (TotalOrderDollars, Customer1PONumber, PODate) " +
                            "VALUES(@TotalOrderDollars, @Customer1PONumber, @PODate)";

                    command.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// Save the detail information from the order to the database
        /// </summary>
        private void SaveCustomer1OrderDetails(Customer1Order order)
        {
            if ((order != null) && (_database.Connection != null) && (_database.Connection.State == ConnectionState.Open))
            {
                foreach (Customer1OrderLineItem lineItem in order.LineItems)
                {
                    using (OleDbCommand command = _database.Connection.CreateCommand())
                    {
                        DbParameter parameter1 = command.CreateParameter();
                        parameter1.ParameterName = "QuantityOrdered";
                        parameter1.DbType = DbType.Int32;
                        parameter1.Value = lineItem.OrderQuantity;
                        command.Parameters.Add(parameter1);

                        DbParameter parameter2 = command.CreateParameter();
                        parameter2.ParameterName = "JSCustomer1ID";
                        parameter2.DbType = DbType.Int32;
                        parameter2.Value = lineItem.JSCustomer1PartNumberID;
                        command.Parameters.Add(parameter2);

                        DbParameter parameter3 = command.CreateParameter();
                        parameter3.ParameterName = "Customer1PONumber";
                        parameter3.DbType = DbType.String;
                        parameter3.Value = order.OrderNumber;
                        command.Parameters.Add(parameter3);

                        DbParameter parameter4 = command.CreateParameter();
                        parameter4.ParameterName = "POPrice";
                        parameter4.DbType = DbType.Int32;
                        parameter4.Value = (lineItem.UnitCost / lineItem.CostPer);
                        command.Parameters.Add(parameter4);

                        DbParameter parameter5 = command.CreateParameter();
                        parameter5.ParameterName = "Customer1PartNumber";
                        parameter5.DbType = DbType.String;
                        parameter5.Value = lineItem.ItemNumber;
                        command.Parameters.Add(parameter5);

                        DbParameter parameter6 = command.CreateParameter();
                        parameter6.ParameterName = "DueByDate";
                        parameter6.DbType = DbType.DateTime;
                        parameter6.Value = lineItem.NeedByDate;
                        command.Parameters.Add(parameter6);

                        command.CommandType = CommandType.Text;
                        command.CommandText = "INSERT INTO [Customer1 Order Details] (QuantityOrdered, JSCustomer1ID, Customer1PONumber, POPrice, Customer1PartNumber, DueByDate) " +
                               "VALUES(@QuantityOrdered, @JSCustomer1ID, @Customer1PONumber, @POPrice, @Customer1PartNumber, @DueByDate)";

                        command.ExecuteNonQuery();
                    }
                }
            }
        }

        #endregion
    }
}
