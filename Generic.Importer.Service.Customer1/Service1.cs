using System;
using System.ServiceProcess;
using System.IO;
using System.Configuration;
using Generic.Importer;
using Generic.Importer.Entities;
using Generic.Importer.Processors;

namespace Generic.Importer.Service.Customer1
{
    public partial class Service1 : ServiceBase
    {
        private Customer1Processor Processor = null;

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            Processor = Customer1Processor.CreateInstance();

            SystemConfiguration config = Processor.ParseConfiguration("ForecastDirectory", "FirmedOrderDirectory");
            Processor.StartProcessing(config, addlDataFile: String.Empty, useSecondaryDataDirectory: true, logToConsole: false);
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
