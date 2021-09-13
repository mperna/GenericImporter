using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using Generic.Importer.Entities;
using Generic.Importer.Managers.Base;

namespace Generic.Importer.Managers
{
    public class Customer2OrderManager : OrderManagerBase
    {
        public Customer2OrderManager(Database database) : base(database) { }

        #region Public Methods

        /// <summary>
        /// Determines if the current order represents data already in the database
        /// </summary>
        public bool Exists(Customer2Order order)
        {
            return base.Exists("Customer2 Orders", "Customer2PONumber", order.PurchaseOrderNumber);
        }

        public void Delete(Customer2Order order)
        {
            base.Delete("Customer2 Orders", "Customer2 Order Details", "Customer2PONumber", order.PurchaseOrderNumber);
        }

        /// <summary>
        /// Retrieves the next revision number by getting the current number from
        /// the sales order record, if it exists, and incrementing it by one.
        /// </summary>
        public int GetOrderNumberRevision(string salesOrderNumber)
        {
            return base.GetOrderNumberRevision("Customer3 Orders", "SalesOrderNumber", "PORevisionLevel", salesOrderNumber);
        }

        public void Save(Customer2Order order)
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

                SaveCustomer2Order(order);
                SaveCustomer2OrderDetails(order);
            }
        }

        #endregion

        #region Private Methods

        private void SaveCustomer2Order(Customer2Order order)
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
                    parameter2.ParameterName = "Customer2PONumber";
                    parameter2.DbType = DbType.String;
                    parameter2.Value = order.PurchaseOrderNumber;
                    command.Parameters.Add(parameter2);

                    DbParameter parameter3 = command.CreateParameter();
                    parameter3.ParameterName = "PODate";
                    parameter3.DbType = DbType.DateTime;
                    parameter3.Value = order.PurchaseOrderDate;
                    command.Parameters.Add(parameter3);

                    DbParameter parameter4 = command.CreateParameter();
                    parameter4.ParameterName = "SalesOrderNumber";
                    parameter4.DbType = DbType.String;
                    parameter4.Value = order.SalesOrderNumber;
                    command.Parameters.Add(parameter4);

                    command.CommandType = CommandType.Text;
                    command.CommandText = "INSERT INTO [Customer2 Orders] (TotalOrderDollars, Customer2PONumber, PODate, SalesOrderNumber) " +
                            "VALUES(@TotalOrderDollars, @Customer2PONumber, @PODate, @SalesOrderNumber)";

                    command.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// Save the detail information from the order to the database
        /// </summary>
        private void SaveCustomer2OrderDetails(Customer2Order order)
        {
            if ((order != null) && (_database.Connection != null) && (_database.Connection.State == ConnectionState.Open))
            {
                foreach (Customer2OrderLineItem lineItem in order.LineItems)
                {
                    using (OleDbCommand command = _database.Connection.CreateCommand())
                    {
                        DbParameter parameter1 = command.CreateParameter();
                        parameter1.ParameterName = "QuantityOrdered";
                        parameter1.DbType = DbType.Int32;
                        parameter1.Value = lineItem.OrderQuantity;
                        command.Parameters.Add(parameter1);

                        DbParameter parameter2 = command.CreateParameter();
                        parameter2.ParameterName = "JSCustomer2ID";
                        parameter2.DbType = DbType.Int32;
                        parameter2.Value = lineItem.JSCustomer2PartNumberID;
                        command.Parameters.Add(parameter2);

                        DbParameter parameter3 = command.CreateParameter();
                        parameter3.ParameterName = "Customer2PONumber";
                        parameter3.DbType = DbType.String;
                        parameter3.Value = order.PurchaseOrderNumber;
                        command.Parameters.Add(parameter3);

                        DbParameter parameter4 = command.CreateParameter();
                        parameter4.ParameterName = "POPrice";
                        parameter4.DbType = DbType.Int32;
                        parameter4.Value = lineItem.ItemCost;
                        command.Parameters.Add(parameter4);

                        DbParameter parameter5 = command.CreateParameter();
                        parameter5.ParameterName = "Customer2PartNumber";
                        parameter5.DbType = DbType.String;
                        parameter5.Value = lineItem.ItemNumber;
                        command.Parameters.Add(parameter5);

                        DbParameter parameter6 = command.CreateParameter();
                        parameter6.ParameterName = "DueByDate";
                        parameter6.DbType = DbType.DateTime;
                        parameter6.Value = lineItem.NeedByDate;
                        command.Parameters.Add(parameter6);

                        command.CommandType = CommandType.Text;
                        command.CommandText = "INSERT INTO [Customer2 Order Details] (QuantityOrdered, JSCustomer2ID, Customer2PONumber, POPrice, Customer2PartNumber, DueByDate) " +
                               "VALUES(@QuantityOrdered, @JSCustomer2ID, @Customer2PONumber, @POPrice, @Customer2PartNumber, @DueByDate)";

                        command.ExecuteNonQuery();
                    }
                }
            }
        }

        #endregion
        
    }
}
