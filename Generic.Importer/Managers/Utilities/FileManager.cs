using System;
using System.IO;
using System.Text.RegularExpressions;

namespace Generic.Importer.Managers.Utilities
{
    public static class FileManager
    {
        #region Public Methods

        /// <summary>
        /// Tests the string provided to determine if it represents a 
        /// valid directory path.
        /// </summary>

        public static bool IsValidDirectoryFormat(string folder)
        {
            bool retVal = false;

            if (String.IsNullOrEmpty(folder) != true)
            {
                Regex r = new Regex(@"^[a-zA-Z]:\\(((?![<>:\/\\|?*]).)+((?<![ .])\\)?)*$");
                retVal = r.IsMatch(folder);

                if (retVal == false)
                {
                    r = new Regex(@"\\\\[a-zA-Z0-9\.\-_]{1,}(\\[a-zA-Z0-9\-_]{1,}){1,}[\$]{0,1}");
                    retVal = r.IsMatch(folder);
                }
            }

            return retVal;
        }

        /// <summary>
        /// Moves the specified file to the desired destination directory, checking
        /// for an existing file and re-naming as needed.
        /// </summary>
        public static string MoveFile(string filePath, string destinationDirectory)
        {
            if ((String.IsNullOrEmpty(filePath) != true) && (String.IsNullOrEmpty(destinationDirectory) != true))
            {
                int count = 0;

                string file = Path.GetFileName(filePath);
                string finalPath = String.Format("{0}\\{1}", destinationDirectory, file);

                while (File.Exists(finalPath) == true)
                {
                    count++;

                    string name = Path.GetFileNameWithoutExtension(filePath);
                    string updatedName = String.Format("{0} ({1})", name, count.ToString());

                    finalPath = String.Format("{0}\\{1}", destinationDirectory, file.Replace(name, updatedName));
                }

                File.Move(filePath, finalPath);
                return finalPath;
            }

            return String.Empty;
        }

        /// <summary>
        /// Removes the specified file from disk.
        /// </summary>
        public static void DeleteFile(string file)
        {
            if (String.IsNullOrEmpty(file) != true)
            {
                File.Delete(file);
            }
        }

        #endregion
    }
}
