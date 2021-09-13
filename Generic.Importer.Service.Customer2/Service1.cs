using System;
using System.ServiceProcess;
using Generic.Importer.Processors;
using Generic.Importer.Entities;

namespace Generic.Importer.Service.Customer2
{
    public partial class Service1 : ServiceBase
    {
        private Customer2Processor Processor = null;

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            Processor = Customer2Processor.CreateInstance();

            SystemConfiguration config = Processor.ParseConfiguration();
            Processor.StartProcessing(config, addlDataFile: String.Empty, useSecondaryDataDirectory: false, logToConsole: false);
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
