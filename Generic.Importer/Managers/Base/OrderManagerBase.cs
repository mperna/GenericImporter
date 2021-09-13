using System;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;

namespace Generic.Importer.Managers.Base
{
    public abstract class OrderManagerBase : ManagerBase
    {
        protected OrderManagerBase(Database database) : base(database) { }

        #region Protected Methods

        /// <summary>
        /// Determines if the current order represents data already in the database
        /// </summary>
        protected bool Exists(string orderTableName, string orderFieldName, string orderNumber)
        {
            bool retVal = false;

            if ((String.IsNullOrEmpty(orderNumber) != true) && (_database.Connection != null) && (_database.Connection.State == ConnectionState.Open))
            {
                //Searching for an existing invoice in the database with this purchase order number.

                using (OleDbCommand command = _database.Connection.CreateCommand())
                {
                    DbParameter parameter = command.CreateParameter();
                    parameter.ParameterName = orderFieldName;
                    parameter.DbType = DbType.String;
                    parameter.Value = orderNumber;

                    command.CommandType = CommandType.Text;
                    command.CommandText = String.Format("SELECT {0} FROM [{1}] WHERE {0} = @{0}", orderFieldName, orderTableName);
                    command.Parameters.Add(parameter);

                    object rev = command.ExecuteScalar();
                    retVal = (rev != null);
                }
            }

            return retVal;
        }

        protected void Delete(string orderTableName, string orderDetailsTableName, string orderFieldName, string orderNumber)
        {
            if ((String.IsNullOrEmpty(orderNumber) != true) && (_database.Connection != null) && (_database.Connection.State == ConnectionState.Open))
            {
                //Delete all existing purchase order detail records

                using (OleDbCommand command = _database.Connection.CreateCommand())
                {
                    DbParameter parameter = command.CreateParameter();
                    parameter.ParameterName = orderFieldName;
                    parameter.DbType = DbType.String;
                    parameter.Value = orderNumber;

                    command.CommandType = CommandType.Text;
                    command.CommandText = String.Format("DELETE FROM [{0}] WHERE {1} = @{1}", orderDetailsTableName, orderFieldName);
                    command.Parameters.Add(parameter);

                    command.ExecuteScalar();
                }

                //Delete the main purchase order record.

                using (OleDbCommand command = _database.Connection.CreateCommand())
                {
                    DbParameter parameter = command.CreateParameter();
                    parameter.ParameterName = orderFieldName;
                    parameter.DbType = DbType.String;
                    parameter.Value = orderNumber;

                    command.CommandType = CommandType.Text;
                    command.CommandText = String.Format("DELETE FROM [{0}] WHERE {1} = @{1}", orderTableName, orderFieldName);
                    command.Parameters.Add(parameter);

                    command.ExecuteScalar();
                }
            }
        }

        /// <summary>
        /// Retrieves the next revision number by getting the current number from
        /// the sales order record, if it exists, and incrementing it by one.
        /// </summary>
        protected int GetOrderNumberRevision(string orderTableName, string orderFieldName, string revisionFieldName, string salesOrderNumber)
        {
            int retVal = 0;

            if ((String.IsNullOrEmpty(salesOrderNumber) != true) && (_database.Connection != null) && (_database.Connection.State == ConnectionState.Open))
            {
                //Searching for an existing invoice in the database with this purchase order number.

                using (OleDbCommand command = _database.Connection.CreateCommand())
                {
                    DbParameter parameter = command.CreateParameter();
                    parameter.ParameterName = orderFieldName;
                    parameter.DbType = DbType.String;
                    parameter.Value = salesOrderNumber;

                    command.CommandType = CommandType.Text;
                    command.CommandText = String.Format("SELECT {0} FROM [{1}] WHERE {2} = @{2}", revisionFieldName, orderTableName, orderFieldName);
                    command.Parameters.Add(parameter);

                    object rev = command.ExecuteScalar();
                    int revision = 0;

                    if (rev != null)
                    {
                        Int32.TryParse(rev.ToString(), out revision);
                        retVal = ((revision != Int32.MinValue) ? revision + 1 : 0);
                    }
                }
            }

            return retVal;
        }

        #endregion
    }
}