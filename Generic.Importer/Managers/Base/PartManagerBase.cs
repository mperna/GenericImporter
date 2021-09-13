using System;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using Generic.Importer.Extensions;

namespace Generic.Importer.Managers.Base
{
    public abstract class PartManagerBase : ManagerBase
    {
        public PartManagerBase(Database database) : base(database) { }

        #region Protected Methods

        /// <summary>
        /// Retrieves the customer1 part number ID for the given part number
        /// </summary>
        protected int GetPartIdentifier(string selectFieldName, string partNumberFieldName, string partNumber)
        {
            if ((String.IsNullOrEmpty(partNumber) != true) && (_database.Connection != null) && (_database.Connection.State == ConnectionState.Open))
            {
                using (OleDbCommand command = _database.Connection.CreateCommand())
                {
                    DbParameter parameter = command.CreateParameter();
                    parameter.ParameterName = partNumberFieldName;
                    parameter.DbType = DbType.String;
                    parameter.Value = partNumber;

                    command.CommandType = CommandType.Text;
                    command.CommandText = String.Format("SELECT {0} FROM [Part Numbers] WHERE {1} = @{2}", selectFieldName, partNumberFieldName, partNumberFieldName);
                    command.Parameters.Add(parameter);

                    object rev = command.ExecuteScalar();
                    return ((rev != null) ? rev.ToString().ToInt32() : 0);
                }
            }

            return 0;
        }

        /// <summary>
        /// Retrieves the customer1 unit cost for the given part number.
        /// </summary>
        protected decimal GetPartUnitCost(string selectFieldName, string partNumberFieldName, string partNumber)
        {
            if ((String.IsNullOrEmpty(partNumber) != true) && (_database.Connection != null) && (_database.Connection.State == ConnectionState.Open))
            {
                using (OleDbCommand command = _database.Connection.CreateCommand())
                {
                    DbParameter parameter = command.CreateParameter();
                    parameter.ParameterName = partNumberFieldName;
                    parameter.DbType = DbType.String;
                    parameter.Value = partNumber;

                    command.CommandType = CommandType.Text;
                    command.CommandText = String.Format("SELECT {0} FROM [Part Numbers] WHERE {1} = @{2}", selectFieldName, partNumberFieldName, partNumberFieldName);
                    command.Parameters.Add(parameter);

                    object rev = command.ExecuteScalar();
                    return ((rev != null) ? rev.ToString().ToDecimal() : 0);
                }
            }

            return 0;
        }

        #endregion
    }
}
