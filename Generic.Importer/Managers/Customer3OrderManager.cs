using Generic.Importer.Entities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Generic.Importer.Managers.Base;

namespace Generic.Importer.Managers
{
    public class Customer3OrderManager : OrderManagerBase
    {
        public Customer3OrderManager(Database database) : base(database) { }

        #region Public Methods

        /// <summary>
        /// Determines if the current invoice represents data already in the database
        /// </summary>
        public bool Exists(Customer3Invoice invoice)
        {
            return base.Exists("Customer3 Orders", "Customer3PONumber", invoice.PurchaseOrder.ToString());
        }

        public void Delete(Customer3Invoice invoice)
        {
            base.Delete("Customer3 Orders", "Customer3 Order Details", "Customer3PONumber", invoice.PurchaseOrder.ToString());
        }

        public void Save(Customer3Invoice invoice)
        {
            if ((invoice != null) && (_database.Connection != null) && (_database.Connection.State == ConnectionState.Open))
            {
                bool recordExists = Exists(invoice);

                //If this purchase order already exists in the database, delete it and its detail records before
                //attempting to re-insert it into the database with the current data.

                if (recordExists == true)
                {
                    Delete(invoice);
                }

                SaveCustomer3Order(invoice);
                SaveCustomer3OrderDetails(invoice);
            }
        }

        /// <summary>
        /// Retrieves the next revision number by getting the current number from
        /// the sales order record, if it exists, and incrementing it by one.
        /// </summary>
        public int GetOrderNumberRevision(string salesOrderNumber)
        {
            return base.GetOrderNumberRevision("Customer3 Orders", "SalesOrderNumber", "PORevisionLevel", salesOrderNumber);
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Save the top level order data from the invoice to the database
        /// </summary>
        private void SaveCustomer3Order(Customer3Invoice invoice)
        {
            if ((invoice != null) && (_database.Connection != null) && (_database.Connection.State == ConnectionState.Open))
            {
                using (OleDbCommand command = _database.Connection.CreateCommand())
                {

                    DbParameter parameter1 = command.CreateParameter();
                    parameter1.ParameterName = "TotalOrderDollars";
                    parameter1.DbType = DbType.Decimal;
                    parameter1.Value = invoice.TotalCost;
                    command.Parameters.Add(parameter1);

                    DbParameter parameter2 = command.CreateParameter();
                    parameter2.ParameterName = "PORevisionLevel";
                    parameter2.DbType = DbType.Int32;
                    parameter2.Value = invoice.Revision;
                    command.Parameters.Add(parameter2);

                    DbParameter parameter3 = command.CreateParameter();
                    parameter3.ParameterName = "SalesOrderNumber";
                    parameter3.DbType = DbType.String;
                    parameter3.Value = invoice.SalesOrderNumber;
                    command.Parameters.Add(parameter3);

                    DbParameter parameter4 = command.CreateParameter();
                    parameter4.ParameterName = "DateEntered";
                    parameter4.DbType = DbType.DateTime;
                    parameter4.Value = DateTime.Now.ToShortTimeString();
                    command.Parameters.Add(parameter4);

                    DbParameter parameter5 = command.CreateParameter();
                    parameter5.ParameterName = "Customer3PONumber";
                    parameter5.DbType = DbType.Int64;
                    parameter5.Value = invoice.PurchaseOrder;
                    command.Parameters.Add(parameter5);

                    DbParameter parameter6 = command.CreateParameter();
                    parameter6.ParameterName = "DueByDate";
                    parameter6.DbType = DbType.DateTime;
                    parameter6.Value = invoice.NeedByDate;
                    command.Parameters.Add(parameter6);

                    DbParameter parameter7 = command.CreateParameter();
                    parameter7.ParameterName = "RouteCode";
                    parameter7.DbType = DbType.String;
                    parameter7.Value = invoice.RouteCode;
                    command.Parameters.Add(parameter7);

                    command.CommandType = CommandType.Text;
                    command.CommandText = "INSERT INTO [Customer3 Orders] (TotalOrderDollars, PORevisionLevel, SalesOrderNumber, DateEntered, Customer3PONumber, DueByDate, [Route Code]) " +
                            "VALUES(@TotalOrderDollars, @PORevisionLevel, @SalesOrderNumber, @DateEntered, @Customer3PONumber, @DueByDate, @RouteCode)";

                    command.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// Save the detail order information from the invoice to the database
        /// </summary>
        private void SaveCustomer3OrderDetails(Customer3Invoice invoice)
        {
            if ((invoice != null) && (_database.Connection != null) && (_database.Connection.State == ConnectionState.Open))
            {
                foreach (Customer3InvoiceLineItem lineItem in invoice.LineItems)
                {
                    using (OleDbCommand command = _database.Connection.CreateCommand())
                    {
                        DbParameter parameter1 = command.CreateParameter();
                        parameter1.ParameterName = "QuantityOrdered";
                        parameter1.DbType = DbType.Int32;
                        parameter1.Value = lineItem.OrderQuantity;
                        command.Parameters.Add(parameter1);

                        DbParameter parameter2 = command.CreateParameter();
                        parameter2.ParameterName = "JSCustomer3ID";
                        parameter2.DbType = DbType.Int32;
                        parameter2.Value = lineItem.Customer3PartNumberID;
                        command.Parameters.Add(parameter2);

                        DbParameter parameter3 = command.CreateParameter();
                        parameter3.ParameterName = "Customer3PONumber";
                        parameter3.DbType = DbType.Int32;
                        parameter3.Value = invoice.PurchaseOrder;
                        command.Parameters.Add(parameter3);

                        DbParameter parameter4 = command.CreateParameter();
                        parameter4.ParameterName = "Customer3PartNumber";
                        parameter4.DbType = DbType.String;
                        parameter4.Value = lineItem.ItemNumber;
                        command.Parameters.Add(parameter4);

                        command.CommandType = CommandType.Text;
                        command.CommandText = "INSERT INTO [Customer3 Order Details] (QuantityOrdered, JSCustomer3ID, Customer3PONumber, Customer3PartNumber) " +
                               "VALUES(@QuantityOrdered, @JSCustomer3ID, @Customer3PONumber, @Customer3PartNumber)";

                        command.ExecuteNonQuery();
                    }
                }
            }
        }

        #endregion
    }
}
