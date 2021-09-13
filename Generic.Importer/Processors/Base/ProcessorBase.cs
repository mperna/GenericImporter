using System;
using System.IO;
using System.Configuration;
using Generic.Importer.Entities;
using Generic.Importer.Managers.Utilities;

namespace Generic.Importer.Processors.Base
{
    public abstract class ProcessorBase
    {
        protected const string _fileFilter = "*.xl*";

        protected string _appUnconfiguredErrorLogName = String.Empty;
        protected string _appDirectoryPath = String.Empty;

        //Standard configuration keys

        private const string _logDirectoryName = "LogDirectory";
        private const string _errorFilesDirectoryName = "ErrorFilesDirectory";
        private const string _processedFilesDirectoryName = "ProcessedFilesDirectory";
        private const string _dataDirectoryConfigName = "DataDirectory";
        private const string _secondaryDataDirectoryConfigName = "SecondaryDataDirectory";

        protected bool _processing = false;

        public ProcessorBase(string defaultErrorLogName)
        {
            string appLocation = System.Reflection.Assembly.GetEntryAssembly().Location;

            _appUnconfiguredErrorLogName = "Errors.log";
            _appDirectoryPath = Path.GetDirectoryName(appLocation);

            DefaultErrorLogName = defaultErrorLogName;
        }

        #region Public Properties

        /// <summary>
        /// Name of the file all error logs are written to
        /// by the LogManager
        /// </summary>
        public string DefaultErrorLogName { get; set; }

        /// <summary>
        /// Indicates if the application is writing log data out to an
        /// active console window.
        /// </summary>
        public bool LogToConsole { get; set; }

        /// <summary>
        /// Directory location where data files are dropped which a FileWatcher
        /// will monitor for activity.
        /// </summary>
        public string DataLocation { get; set; }

        /// <summary>
        /// Directory location where secondary data files are dropped which a FileWatcher
        /// will monitor for activity.
        /// </summary>
        public string SecondaryDataLocation { get; set; }

        /// <summary>
        /// Directory location where all logs are written during file
        /// and application processing.
        /// </summary>
        public string LogLocation { get; set; }

        /// <summary>
        /// Directory location of all data files which encounter errors 
        /// during processing
        /// </summary>
        public string ErrorsLocation { get; set; }

        /// <summary>
        /// Directory location of all data fiels which successfully complete
        /// their processing cycle.
        /// </summary>
        public string ProcessedLocation { get; set; }

        /// <summary>
        /// Component which monitors the indicated drop folder for
        /// data files.
        /// </summary>
        public FileSystemWatcher FileWatcher1 { get; set; }

        /// <summary>
        /// Component which monitors the indicated drop folder for
        /// secondary data files.
        /// </summary>
        public FileSystemWatcher FileWatcher2 { get; set; }

        /// <summary>
        /// Delegate method used to process files dropped into the primary data folder
        /// provided by the associated StartProcessing override.
        /// </summary>
        public Action<string, string> ProcessPrimaryFile { get; set; }

        /// <summary>
        /// Delegate method used to process files dropped into the secondary data folder
        /// provided by the associated StartProcessing override.
        /// </summary>
        public Action<string, string> ProcessSecondaryFile { get; set; }

        #endregion

        #region Public Methods

        /// <summary>
        /// Extracts all standard configuration values from the local app.config file for use by
        /// the processor classes.
        /// </summary>
        /// <param name="customDataDirectoryConfigName">Custom configuration name for the data directory</param>
        /// <param name="customSecondaryDataDirectoryConfigName">Custom configuration name for the secondary data directory</param>
        public virtual SystemConfiguration ParseConfiguration(string customDataDirectoryConfigName = "", string customSecondaryDataDirectoryConfigName = "")
        {
            var retVal = new SystemConfiguration();

            string dataDirectoryName = String.IsNullOrEmpty(customDataDirectoryConfigName) ? _dataDirectoryConfigName : customDataDirectoryConfigName;
            string secondaryDataDirectoryName = String.IsNullOrEmpty(customSecondaryDataDirectoryConfigName) ? _secondaryDataDirectoryConfigName : customSecondaryDataDirectoryConfigName;

            //Data Directories

            retVal.DataDirectory = ConfigurationManager.AppSettings[dataDirectoryName].ToString();

            if (ConfigurationManager.AppSettings[secondaryDataDirectoryName] != null)
            {
                retVal.SecondaryDataDirectory = ConfigurationManager.AppSettings[secondaryDataDirectoryName].ToString();
            }
            
            //Supporting Directories

            retVal.LogFilesDirectory = ConfigurationManager.AppSettings[_logDirectoryName].ToString();
            retVal.ErrorFilesDirectory = ConfigurationManager.AppSettings[_errorFilesDirectoryName].ToString();
            retVal.ProcessedFilesDirectory = ConfigurationManager.AppSettings[_processedFilesDirectoryName].ToString();
            
            return retVal;
        }

        /// <summary>
        /// Method which starts the process of analyzing and parsing data files for the implementation
        /// class; setting up the FileSystemWatcher as well as validating all start up properties.
        /// </summary>
        /// <param name="appConfig">Stanard configuration values found within the local app.config</param>
        /// <param name="addlDataFile">An optional data file to aid in the processing of files dropped within either of the data directories</param>
        /// <param name="useSecondaryDataDirectory">Indicates the processor should attempt to use the secondary data directory, if configured</param>
        /// <param name="logToConsole">Indicates if the application is running in a console window and should display all messags in that interface</param>
        /// <returns>Boolean value indicating if process start-up was successful</returns>
        public virtual bool StartProcessing(SystemConfiguration appConfig, string addlDataFile = "", bool useSecondaryDataDirectory = false, bool logToConsole = false)
        {
            bool stopProcessing = false;

            //Initializing the flag which indicates if the application is processing files
            //through a console window and display errors within that interface.

            LogToConsole = logToConsole;

            if (FileManager.IsValidDirectoryFormat(appConfig.DataDirectory) != true)
            {
                stopProcessing = true;

                string message = String.Format("The configured data directory path is invalid: {0}", appConfig.DataDirectory);
                LogManager.WriteLog(_appUnconfiguredErrorLogName, _appDirectoryPath, message, isError: true, mirrorToConsole: logToConsole);
            }
            else if (Directory.Exists(appConfig.DataDirectory) != true)
            {
                Directory.CreateDirectory(appConfig.DataDirectory);
            }

            DataLocation = appConfig.DataDirectory;

            if (useSecondaryDataDirectory == true)
            {
                if (FileManager.IsValidDirectoryFormat(appConfig.SecondaryDataDirectory) != true)
                {
                    stopProcessing = true;

                    string message = String.Format("The configured secondary data directory path is invalid: {0}", appConfig.SecondaryDataDirectory);
                    LogManager.WriteLog(_appUnconfiguredErrorLogName, _appDirectoryPath, message, isError: true, mirrorToConsole: logToConsole);
                }
                else if (Directory.Exists(appConfig.SecondaryDataDirectory) != true)
                {
                    Directory.CreateDirectory(appConfig.SecondaryDataDirectory);
                }

                SecondaryDataLocation = appConfig.SecondaryDataDirectory;
            }

            if (FileManager.IsValidDirectoryFormat(appConfig.LogFilesDirectory) != true)
            {
                stopProcessing = true;

                string message = String.Format("The configured log directory path is invalid: {0}", appConfig.LogFilesDirectory);
                LogManager.WriteLog(_appUnconfiguredErrorLogName, _appDirectoryPath, message, isError: true, mirrorToConsole: logToConsole);
            }
            else if (Directory.Exists(appConfig.LogFilesDirectory) != true)
            {
                Directory.CreateDirectory(appConfig.LogFilesDirectory);
            }

            LogLocation = appConfig.LogFilesDirectory;

            if (FileManager.IsValidDirectoryFormat(appConfig.ProcessedFilesDirectory) != true)
            {
                stopProcessing = true;

                string message = String.Format("The configured processed files directory path is invalid: {0}", appConfig.ProcessedFilesDirectory);
                LogManager.WriteLog(_appUnconfiguredErrorLogName, _appDirectoryPath, message, isError: true, mirrorToConsole: logToConsole);
            }
            else if (Directory.Exists(appConfig.ProcessedFilesDirectory) != true)
            {
                Directory.CreateDirectory(appConfig.ProcessedFilesDirectory);
            }

            ProcessedLocation = appConfig.ProcessedFilesDirectory;

            if (FileManager.IsValidDirectoryFormat(appConfig.ErrorFilesDirectory) != true)
            {
                stopProcessing = true;

                string message = String.Format("The configured error files directory path is invalid: {0}", appConfig.ErrorFilesDirectory);
                LogManager.WriteLog(_appUnconfiguredErrorLogName, _appDirectoryPath, message, isError: true, mirrorToConsole: logToConsole);
            }
            else if (Directory.Exists(appConfig.ErrorFilesDirectory) != true)
            {
                Directory.CreateDirectory(appConfig.ErrorFilesDirectory);
            }

            ErrorsLocation = appConfig.ErrorFilesDirectory;

            if ((stopProcessing != true) && (useSecondaryDataDirectory == true))
            {
                //Initializing the file system watcher to monitor for new secondary data files created within 
                //the defined folder location.

                FileWatcher2 = new FileSystemWatcher(SecondaryDataLocation, _fileFilter);
                FileWatcher2.Created += FileWatcher2_Created;
                FileWatcher2.EnableRaisingEvents = false;
            }

            if (stopProcessing != true)
            {
                //Initializing the file system watcher to monitor for new data files created within 
                //the defined folder location.

                FileWatcher1 = new FileSystemWatcher(DataLocation, _fileFilter);
                FileWatcher1.Created += FileWatcher1_Created;
                FileWatcher1.EnableRaisingEvents = false;

                return true;
            }

            return false;
        }

        /// <summary>
        /// Method which stops all automated processes which monitor and process data files in
        /// the configured folder location.
        /// </summary>
        public virtual void StopProcessing()
        {
            if (FileWatcher1 != null)
            {
                FileWatcher1.EnableRaisingEvents = false;
            }

            if (FileWatcher2 != null)
            {
                FileWatcher2.EnableRaisingEvents = false;
            }
        }

        /// <summary>
        /// Method which processes a top-level file processing error by logging it to the appropriate
        /// file/directory or displaying it, as needed, on the console.
        /// </summary>
        /// <param name="ex">Exception thrown</param>
        /// <param name="fileName">Name of the file being processed</param>
        /// <param name="filePath">Directory path of the file being processed</param>
        public virtual void LogProcessingError(Exception ex, string fileName, string filePath)
        {
            if (LogToConsole == true)
            {
                Console.ForegroundColor = ConsoleColor.Red;
            }

            //Add all appropriate log entries to document that this file encountered an error during
            //processing prior to moving it to the directory for processing errors.

            LogManager.WriteLog(DefaultErrorLogName, LogLocation, String.Format("Parsing failed for file '{0}'; the error message follows on the next line.", fileName), isError: true, mirrorToConsole: LogToConsole);
            LogManager.WriteLog(DefaultErrorLogName, LogLocation, ex.Message, isError: true, stackTrace: ex.StackTrace, mirrorToConsole: LogToConsole);

            try
            {
                //Move the file to the error files location and delete the 
                //remaining file from the drop location.

                string movedFile = FileManager.MoveFile(filePath, ErrorsLocation);
                LogManager.WriteLog(DefaultErrorLogName, LogLocation, String.Format("Moving '{0}' to '{1}' for storage", filePath.ToLower(), movedFile.ToLower()), mirrorToConsole: LogToConsole);

                //Deleting the original file from the file processing directory.
                
                if (File.Exists(filePath) == true)
                {
                    FileManager.DeleteFile(filePath);
                }
            }
            catch (Exception)
            {
                LogManager.WriteLog(DefaultErrorLogName, LogLocation, String.Format("An error occurred attempting to move {0} to the error directory following a processing error.", filePath.ToLower()), isError: true, mirrorToConsole: LogToConsole);
            }

            if (LogToConsole == true)
            {
                Console.ForegroundColor = ConsoleColor.White;
            }
        }

        #endregion

        #region Private Methods

        private void FileWatcher1_Created(object sender, FileSystemEventArgs e)
        {
            while (_processing == true)
            {
                System.Threading.Thread.Sleep(5000);
            }

            _processing = true;

            ProcessPrimaryFile(e.Name, e.FullPath);
            System.Threading.Thread.Sleep(2000);

            _processing = false;
        }

        private void FileWatcher2_Created(object sender, FileSystemEventArgs e)
        {
            while (_processing == true)
            {
                System.Threading.Thread.Sleep(5000);
            }

            _processing = true;

            ProcessSecondaryFile(e.Name, e.FullPath);
            System.Threading.Thread.Sleep(2000);

            _processing = false;
        }

        #endregion
    }
}
