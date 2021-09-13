using System;
using System.ServiceProcess;
using System.IO;
using System.Configuration;
using Generic.Importer;
using Generic.Importer.Processors;
using Generic.Importer.Entities;

namespace Generic.Importer.Service.Customer3
{
    public partial class Service1 : ServiceBase
    {
        private Customer3Processor Processor = null;

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            Processor = Customer3Processor.CreateInstance();
            string cheatSheet = ConfigurationManager.AppSettings["CheatSheet"].ToString();

            SystemConfiguration config = Processor.ParseConfiguration();
            Processor.StartProcessing(config, addlDataFile: cheatSheet, useSecondaryDataDirectory: false, logToConsole: false);
        }

        protected override void OnStop()
        {
            if (Processor != null)
            {
                Processor.StopProcessing();
            }
        }
    }
}
