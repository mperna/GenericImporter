using System;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.IO;
using Generic.Importer.Entities;
using Generic.Importer.Extensions;

namespace Generic.Importer
{
    public class Database : IDisposable
    {
        public OleDbConnection Connection { get; set; }
        public string FileLocation { get; set; }

        public Database()
        {
            string fileLocation = ConfigurationManager.AppSettings["Database"];
            if ((String.IsNullOrEmpty(fileLocation) == true) || (File.Exists(fileLocation) != true))
            {
                throw new Exception("No database found for processing; check the configuration file for the proper file location.");
            }

            FileLocation = fileLocation;
        }

        /// <summary>
        /// Creates and opens a connection to the Access database defined
        /// within the app.config.
        /// </summary>
        public void OpenConnection()
        {
            Connection = null;

            if (String.IsNullOrEmpty(FileLocation) != true)
            {
                Connection = new OleDbConnection();
                Connection.ConnectionString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Persist Security Info=False;", FileLocation);

                if (Connection == null)
                {
                    throw new Exception("Unable to create an oledb connection to the database provided.");
                }

                Connection.Open();
            }
        }

        public void CloseConnection()
        { 
            if ((Connection != null) && (Connection.State == ConnectionState.Open))
            {
                Connection.Close();
            }

            Connection.Dispose();
        }

        public void Dispose()
        {
            if ((Connection != null) & (Connection.State == ConnectionState.Open))
            {
                Connection.Close();
                Connection = null;
            }
        }
    }
}
