
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Generic.Importer.Entities;
using Generic.Importer.Extensions;
using Generic.Importer.Managers;
using Generic.Importer.Managers.Utilities;
using Generic.Importer.Processors.Base;


namespace Generic.Importer.Processors
{
    public class Customer1Processor : ProcessorBase
    {
        private struct DataOrdinals
        {
            public const int Order_Excel_Ordinal_OrderNumber = 0;
            public const int Order_Excel_Ordinal_LineNumber = 1;
            public const int Order_Excel_Ordinal_ItemNumber = 2;
            public const int Order_Excel_Ordinal_Quantity = 8;
            public const int Order_Excel_Ordinal_NeedByDate = 12;
            public const int Order_Excel_Ordinal_UnitPrice = 17;
            public const int Order_Excel_Ordinal_CostPer = 18;
            public const int Order_Excel_Ordinal_ImportDate = 36;
            public const int Order_Excel_NeededColumns = 36;
        }

        private const string _pricingLogName = "Customer1Pricing";
        private const string _logName = "Customer1Log";

        public Customer1Processor() : base(_logName)
        {
            //Do Nothing
        }

        #region Public Methods

        public override bool StartProcessing(SystemConfiguration appConfig, string addlDataFile = "", bool useSecondaryDataDirectory = false, bool logToConsole = false)
        {
            //Validation and initialization all standard file folders as well as the file system watcher within
            //the base class in order to monitor for new data files within the defined folder location.

            bool success = base.StartProcessing(appConfig, addlDataFile, useSecondaryDataDirectory, logToConsole);
            bool stopProcessing = false;

            if ((stopProcessing != true) && (success == true))
            {
                try
                {
                    //Before initializing the file watchers, process any existing forecast or firmed order files 
                    //which are currently waiting to be processed.

                    var existingForcastFiles = Directory.GetFiles(DataLocation, _fileFilter).OrderByDescending(d => new FileInfo(d).CreationTime);
                    var existingFirmedOrderFiles = Directory.GetFiles(SecondaryDataLocation, _fileFilter).OrderByDescending(d => new FileInfo(d).CreationTime);

                    foreach (string file in existingFirmedOrderFiles)
                    {
                        ProcessFirmedOrder(Path.GetFileName(file), file);
                    }

                    foreach (string file in existingForcastFiles)
                    {
                        ProcessForecast(Path.GetFileName(file), file);
                    }

                    //Initializing the delegate properties which will enable the FileWatchers to process
                    //data files using methods defined in this class.

                    base.ProcessPrimaryFile = ProcessForecast;
                    base.ProcessSecondaryFile = ProcessFirmedOrder;

                    //Start processing all new data files dropped within the DataLocation folder.

                    FileWatcher1.EnableRaisingEvents = true;
                    FileWatcher2.EnableRaisingEvents = true;
                }
                catch (Exception ex)
                {
                    LogManager.WriteLog(DefaultErrorLogName, LogLocation, ex.Message, isError: true, stackTrace: ex.StackTrace, useDatedLogs: true, mirrorToConsole: LogToConsole);
                }
            }

            return ((stopProcessing != true) && (success == true));
        }

        public void ProcessForecast(string name, string filePath)
        {
            try
            {
                if (LogToConsole == true)
                {
                    Console.WriteLine(String.Format("File Found: {0}", name));
                }

                //Establishing a connection to the Access database defined within the
                //app.config file and saving the parsed data to it.

                using (Database database = new Database())
                {
                    database.OpenConnection();

                    //Parsing data from the physical file prior to storing it in the
                    //database.

                    List<Customer1Order> orders = ParseExcelOrder(database, filePath);
                    Customer1OrderManager orderManager = new Customer1OrderManager(database);

                    //Purging all forecasts ahead of re-importing forecast data using the parsed
                    //data files.

                    orderManager.PurgeForecasts();

                    //Save all invoice information to the order / order detail tables in
                    //the Customer3 database.

                    foreach (Customer1Order order in orders)
                    {
                        try
                        {
                            //Searching each line item for discrepencies between the stated part cost on the 
                            //PO and the Customer1 part cost in the database before saving.

                            foreach (Customer1OrderLineItem lineItem in order.LineItems)
                            {
                                if (lineItem.UnitCost.Equals(lineItem.Customer1UnitCost) != true)
                                {
                                    LogManager.WriteLog(_pricingLogName, LogLocation, String.Format("Part number {0} has a forecast unit cost of {1} on order {2} and a Customer1 unit cost of {3}.",
                                        lineItem.ItemNumber, lineItem.UnitCost, order.OrderNumber, lineItem.Customer1UnitCost), stackTrace: String.Empty);
                                }
                            }

                            orderManager.Save(order);

                            //Add all appropriate log entries to document that this file was processed prior to
                            //moving it to the processed files location.

                            LogManager.WriteLog(DefaultErrorLogName, LogLocation, String.Format("Processing completed for forecast '{0}' with {1} line items as part of file {2}", order.OrderNumber, order.LineItems.Count, name), mirrorToConsole: LogToConsole);
                        }
                        catch (Exception ex)
                        {
                            if (LogToConsole == true)
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                            }

                            //Add all appropriate log entries to document that an error was encountered while processing
                            //invoice data within this file.

                            LogManager.WriteLog(DefaultErrorLogName, LogLocation, String.Format("An error was encountered while attempting to save forecast '{0}' as part of file {1}; the error is on the next line.", order.OrderNumber, name), isError: true, mirrorToConsole: LogToConsole);
                            LogManager.WriteLog(DefaultErrorLogName, LogLocation, ex.Message, isError: true, stackTrace: ex.StackTrace, mirrorToConsole: LogToConsole);

                            if (LogToConsole == true)
                            {
                                Console.ForegroundColor = ConsoleColor.White;
                            }
                        }
                    }

                    string movedFile = FileManager.MoveFile(filePath, ProcessedLocation);
                    LogManager.WriteLog(DefaultErrorLogName, LogLocation, String.Format("Processing complete for file {0}; moving to {1} for storage.", filePath.ToLower(), movedFile.ToLower()), mirrorToConsole: LogToConsole);

                    database.CloseConnection();
                }

                if (File.Exists(filePath) == true)
                {
                    FileManager.DeleteFile(filePath);
                }
            }
            catch (Exception ex)
            {
                LogProcessingError(ex, name, filePath);
            }
        }

        public void ProcessFirmedOrder(string name, string filePath)
        {
            try
            {
                //Establishing a connection to the Access database defined within the
                //app.config file and saving the parsed data to it.

                using (Database database = new Database())
                {
                    database.OpenConnection();

                    //Parsing data from the physical file prior to storing it in the
                    //database.

                    List<Customer1Order> orders = ParseExcelOrder(database, filePath);
                    Customer1OrderManager orderManager = new Customer1OrderManager(database);

                    //Save all invoice information to the order / order detail tables in
                    //the Customer3 database.

                    foreach (Customer1Order order in orders)
                    {
                        try
                        {
                            //Searching each line item for discrepencies between the stated part cost on the 
                            //PO and the Customer1 part cost in the database before saving.

                            foreach (Customer1OrderLineItem lineItem in order.LineItems)
                            {
                                if (lineItem.UnitCost.Equals(lineItem.Customer1UnitCost) != true)
                                {
                                    LogManager.WriteLog(_pricingLogName, LogLocation, String.Format("Part number {0} has a firmed order unit cost of {1} on order {2} and a Customer1 unit cost of {3}.",
                                        lineItem.ItemNumber, lineItem.UnitCost, order.OrderNumber, lineItem.Customer1UnitCost), stackTrace: String.Empty);
                                }
                            }

                            orderManager.Save(order);

                            //Add all appropriate log entries to document that this file was processed prior to
                            //moving it to the processed files location.

                            LogManager.WriteLog(DefaultErrorLogName, LogLocation, String.Format("Processing completed for firmed order '{0}' with {1} line items as part of file {2}", order.OrderNumber, order.LineItems.Count, name), mirrorToConsole: LogToConsole);
                        }
                        catch (Exception ex)
                        {
                            if (LogToConsole == true)
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                            }

                            //Add all appropriate log entries to document that an error was encountered while processing
                            //invoice data within this file.

                            LogManager.WriteLog(DefaultErrorLogName, LogLocation, String.Format("An error was encountered while attempting to save firmed order '{0}' as part of file {1}; the error is on the next line.", order.OrderNumber, name), isError: true, mirrorToConsole: LogToConsole);
                            LogManager.WriteLog(DefaultErrorLogName, LogLocation, ex.Message, isError: true, stackTrace: ex.StackTrace, mirrorToConsole: LogToConsole);

                            if (LogToConsole == true)
                            {
                                Console.ForegroundColor = ConsoleColor.White;
                            }
                        }
                    }

                    string movedFile = FileManager.MoveFile(filePath, ProcessedLocation);
                    LogManager.WriteLog(DefaultErrorLogName, LogLocation, String.Format("Processing complete for file {0}; moving to {1} for storage.", filePath.ToLower(), movedFile.ToLower()), mirrorToConsole: LogToConsole);

                    database.CloseConnection();
                }

                if (File.Exists(filePath) == true)
                {
                    FileManager.DeleteFile(filePath);
                }
            }
            catch (Exception ex)
            {
                LogProcessingError(ex, name, filePath);
            }
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Parses the excel file provided using the current database connection for any 
        /// supplementary look-ups required.
        /// </summary>
        private List<Customer1Order> ParseExcelOrder(Database database, string filePath)
        {
            List<Customer1Order> retVal = new List<Customer1Order>();
            Exception caught = null;

            string ext = Path.GetExtension(filePath);

            IWorkbook book = null;
            ISheet sheet = null;

            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    if (ext.ToLower().Equals(".xls") == true)
                    {
                        book = new HSSFWorkbook(fs);
                    }
                    else if (ext.ToLower().Equals(".xlsx") == true)
                    {
                        book = new XSSFWorkbook(fs);
                    }

                    if (book == null)
                    {
                        throw new Exception("Unable to initialize the library necessary to read the Excel file provided.");
                    }

                    sheet = book.GetSheetAt(0);
                    var enumerator = sheet.GetRowEnumerator();

                    if (enumerator == null)
                    {
                        throw new Exception("Unable to establish the enumerator required to parse the provided Excel file.");
                    }

                    while (enumerator.MoveNext() == true)
                    {
                        IRow lineItem = (IRow)enumerator.Current;

                        if (lineItem == null)
                        {
                            throw new Exception("Unexpected null row data found in Excel file provided.");
                        }

                        if (lineItem.Cells.Count < DataOrdinals.Order_Excel_NeededColumns)
                        {
                            throw new Exception("The number of columns expected in the Excel spreadsheet does not equal the number found.");
                        }

                        //Check if this is the header row; if so, by-pass it and begin processing
                        //the actual sheet data in subsequent rows.

                        if (lineItem.Cells[DataOrdinals.Order_Excel_Ordinal_OrderNumber].CellType != CellType.String)
                        {
                            double orderNumber = lineItem.Cells[DataOrdinals.Order_Excel_Ordinal_OrderNumber].NumericCellValue;
                            Customer1Order order = retVal.Find(x => x.OrderNumber.Equals(orderNumber));

                            //If this order is not yet within our list of forecasts then this is the first time we've encountered
                            //it in the excel file; create a new record, import its values and process all of its line items.

                            if (order == null)
                            {
                                order = new Customer1Order();

                                order.Revision = 0;
                                order.OrderNumber = orderNumber;
                                order.ImportDate = lineItem.Cells[DataOrdinals.Order_Excel_Ordinal_ImportDate].GetStringValue();
                               
                                //This first record will also contain data for the first line item in the forecast; parse it out
                                //and add the line item record to the invoice.

                                Customer1OrderLineItem orderlineItem = ParseExcelOrderLineItem(database, lineItem);

                                //Setting the total cost based on this initial line item and adding the corresponding line item
                                //to the order's internal collection.

                                order.TotalCost = (orderlineItem.OrderQuantity * (orderlineItem.UnitCost / orderlineItem.CostPer));
                                order.TotalCustomer1Cost = (orderlineItem.OrderQuantity * orderlineItem.Customer1UnitCost);

                                order.LineItems.Add(orderlineItem);

                                //Add this new order to the results collection

                                retVal.Add(order);
                            }
                            else
                            {
                                //If this order number already exists within our list of orders then this is another line item on that
                                //order; create a new line item record, process its values and add it to the invoice.

                                Customer1OrderLineItem orderlineItem = ParseExcelOrderLineItem(database, lineItem);

                                //Updating the total cost for this order based on this values in this line item and adding
                                //the line item to the orders internal collection.

                                order.TotalCost += (orderlineItem.OrderQuantity * (orderlineItem.UnitCost / orderlineItem.CostPer));
                                order.TotalCustomer1Cost += (orderlineItem.OrderQuantity * orderlineItem.Customer1UnitCost);

                                order.LineItems.Add(orderlineItem);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //If an exception is thrown, catch it here so that the finalization
                //logic can execute before re-throwing it for the calling routine to handle.

                caught = ex;
            }
            finally
            {
                if (book != null)
                {
                    book.Close();
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();

                book = null;
                sheet = null;
            }

            if (caught != null)
            {
                throw caught;
            }

            return retVal;
        }

        /// <summary>
        /// Parses out the line item data from the given excel row
        /// </summary>
        private Customer1OrderLineItem ParseExcelOrderLineItem(Database database, IRow row)
        {
            Customer1OrderLineItem retVal = new Customer1OrderLineItem();

            retVal.ItemNumber = row.Cells[DataOrdinals.Order_Excel_Ordinal_ItemNumber].GetStringValue();
            retVal.LineNumber = row.Cells[DataOrdinals.Order_Excel_Ordinal_LineNumber].NumericCellValue;
            retVal.CostPer = row.Cells[DataOrdinals.Order_Excel_Ordinal_CostPer].NumericCellValue;
            retVal.UnitCost = row.Cells[DataOrdinals.Order_Excel_Ordinal_UnitPrice].NumericCellValue;
            retVal.OrderQuantity = row.Cells[DataOrdinals.Order_Excel_Ordinal_Quantity].NumericCellValue;
            retVal.NeedByDate = row.Cells[DataOrdinals.Order_Excel_Ordinal_NeedByDate].GetStringValue();

            Customer1PartManager partManager = new Customer1PartManager(database);
            
            retVal.ItemNumber = retVal.ItemNumber.Trim().Substring(8);
            retVal.Customer1UnitCost = (double)partManager.GetPartUnitCost(retVal.ItemNumber);
            retVal.JSCustomer1PartNumberID = partManager.GetPartIdentifier(retVal.ItemNumber);
            
            return retVal;
        }

        #endregion

        public static Customer1Processor CreateInstance()
        {
            return new Customer1Processor();
        }
    }
}
