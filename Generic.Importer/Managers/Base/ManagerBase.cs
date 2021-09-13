using System;

namespace Generic.Importer.Managers.Base
{
    public abstract class ManagerBase
    {
        protected Database _database = null;

        public ManagerBase(Database database)
        {
            if (database == null)
            {
                throw new Exception("The database provided is not initialized for transactions.");
            }
            
            _database = database;
        }
    }
}
