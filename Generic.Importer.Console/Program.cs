using System;
using System.Configuration;
using Generic.Importer.Entities;
using Generic.Importer.Processors;
using Generic.Importer.Processors.Base;


namespace Generic.Importer.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            ProcessorBase processor = null;

            if (args.Length >= 1)
            {
                if (String.Equals(args[0].ToLower(), "-k") == true)
                {
                    //Processing standard Customer3 invoice files

                    processor = Customer3Processor.CreateInstance();
                    string cheatSheet = ConfigurationManager.AppSettings["CheatSheet"].ToString();

                    SystemConfiguration config = processor.ParseConfiguration();
                    processor.StartProcessing(config, addlDataFile: cheatSheet, useSecondaryDataDirectory: false, logToConsole: true);
                }
                else if (String.Equals(args[0].ToLower(), "-h") == true)
                {
                    //Processing Customer1 input files

                    processor = Customer1Processor.CreateInstance();

                    SystemConfiguration config = processor.ParseConfiguration("ForecastDirectory", "FirmedOrderDirectory");
                    processor.StartProcessing(config, addlDataFile: String.Empty, useSecondaryDataDirectory: true, logToConsole: true);

                }
                else if (String.Equals(args[0].ToLower(), "-hs") == true)
                {
                    //Processing Human Scale input files

                    processor = Customer2Processor.CreateInstance();

                    SystemConfiguration config = processor.ParseConfiguration();
                    processor.StartProcessing(config, addlDataFile: String.Empty, useSecondaryDataDirectory: false, logToConsole: true);
                }
            }

            //Waiting for user input to continue

            System.Console.ReadLine();

            //Terminating all file processsing.

            processor.StopProcessing();
        }
    }
}
