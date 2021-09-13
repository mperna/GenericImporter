using System;
using System.IO;
using System.Collections.Generic;

namespace Generic.Importer.Managers.Utilities
{
    public static class LogManager
    {
        #region Public Methods

        /// <summary>
        /// Writes a log entry into a specified log file on disk
        /// </summary>
        /// If a log file does not exist at the log location, this method will create one
        /// </remarks>
        public static void WriteLog(string logLocation, string message, bool isError = false, string stackTrace = "", bool useDatedLogs = true, bool mirrorToConsole = true)
        {
            WriteLog("Log", logLocation, message, isError, stackTrace, useDatedLogs, mirrorToConsole);
        }

        public static void WriteLog(string logName, string logLocation, string message, bool isError = false, string stackTrace = "", bool useDatedLogs = true, bool mirrorToConsole = true)
        {
            if ((String.IsNullOrEmpty(logLocation) != true) && (String.IsNullOrEmpty(message) != true) && (Directory.Exists(logLocation) == true))
            {
                FileStream file = null;
                string logFullNameLocation = String.Empty;

                if (useDatedLogs == true)
                {
                    //All log files are created using the current date as a portion of the log name.

                    logName = String.Format("{0}-{1}.txt", logName, DateTime.Now.ToString("MMddyyyy"));
                }

                logFullNameLocation = String.Format("{0}\\{1}", logLocation, logName);

                if (File.Exists(logFullNameLocation) != true)
                {
                    file = File.Create(logFullNameLocation);
                    file.Close();
                }

                //Add log entry to the log file, appending a date/time stamp to the message

                string prefix = ((isError == true) ? "Error" : "Status");
                List<string> content = new List<string>() { String.Format("{0}: {1} - {2}", prefix, DateTime.Now.ToLongTimeString(), message) };

                if (String.IsNullOrEmpty(stackTrace) != true)
                {
                    content.Add(String.Format("{0} {1}", DateTime.Now.ToLongTimeString(), stackTrace));
                }

                File.AppendAllLines(logFullNameLocation, content);

                if (mirrorToConsole == true)
                {
                    Console.WriteLine(message);
                }
            }
        }

        #endregion
    }
}
