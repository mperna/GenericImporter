using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Data;
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
    public class Customer3Processor : ProcessorBase
    {
        private struct DataOrdinals
        {
            public const int Invoice_Excel_Ordinal_Supplier = 0;
            public const int Invoice_Excel_Ordinal_PurchaseOrder = 1;
            public const int Invoice_Excel_Ordinal_Revision = 3;
            public const int Invoice_Excel_Ordinal_OrderDate = 4;
            public const int Invoice_Excel_Ordinal_OrderNumber = 6;
            public const int Invoice_Excel_Ordinal_HeaderStatus = 8;
            public const int Invoice_Excel_Ordinal_LineNumber = 10;
            public const int Invoice_Excel_Ordinal_Quantity = 11;
            public const int Invoice_Excel_Ordinal_UnitPrice = 14;
            public const int Invoice_Excel_Ordinal_ItemStatus = 16;
            public const int Invoice_Excel_Ordinal_ItemNumber = 17;
            public const int Invoice_Excel_Ordinal_RouteCode = 19;
            public const int Invoice_Excel_Ordinal_NeedByDate = 20;
            public const int Invoice_Excel_NeededColumns = 21;

            public const int Substitution_Excel_StartsOnRow = 3;
            public const int Substitution_Excel_ShippablePackItemNumber = 1;
            public const int Substitution_Excel_ShippablePackItemNumberDescription = 2;
            public const int Substitution_Excel_ImportKitItemNumberQuantity = 4;
            public const int Substitution_Excel_ImportKitItemNumber = 5;
            public const int Substitution_Excel_ImportKitItemNumberDescription = 6;
            public const int Substitution_Excel_CrossbarsItemNumberQuantity = 8;
            public const int Substitution_Excel_CrossbarsItemNumber = 9;
            public const int Substitution_Excel_CrossbarsItemNumberDescription = 10;
        }

        private const string _pricingLogName = "Customer3Pricing";
        private const string _logName = "Customer3Log";

        public Customer3Processor() : base(_logName)
        {
            //Do Nothing
        }

        #region Private Properties

        private string SubstitutionLocation { get; set; }

        #endregion

        #region Public Methods

        public override bool StartProcessing(SystemConfiguration appConfig, string addlDataFile = "", bool useSecondaryDataDirectory = false, bool logToConsole = false)
        {
            //Validation and initialization all standard file folders as well as the file system watcher within
            //the base class in order to monitor for new data files within the defined folder location.

            bool success = base.StartProcessing(appConfig, addlDataFile, useSecondaryDataDirectory, logToConsole);
            bool stopProcessing = false;

            //If a cheat sheet is provided, ensure that it exists and, if not, throw an
            //exception and stop all processing.

            SubstitutionLocation = addlDataFile;
            if ((String.IsNullOrEmpty(SubstitutionLocation) != true) && (File.Exists(SubstitutionLocation) != true))
            {
                stopProcessing = true;

                string message = String.Format("The configured substitution sheet is invalid or does not appear to exist on disk: {0}", addlDataFile);
                LogManager.WriteLog(_appUnconfiguredErrorLogName, _appDirectoryPath, message, isError: true, mirrorToConsole: logToConsole);
            }

            if ((stopProcessing != true) && (success == true))
            {
                try
                {
                    //Before initializing the file watcher, parse any existing files which are currently
                    //waiting to be processed.

                    string[] existingFiles = Directory.GetFiles(DataLocation, _fileFilter);

                    if (existingFiles.Count() > 0)
                    {
                        foreach (string file in existingFiles)
                        {
                            ProcessInvoice(Path.GetFileName(file), file);
                        }
                    }

                    //Initializing the delegate properties which will enable the FileWatchers to process
                    //data files using methods defined in this class.

                    base.ProcessPrimaryFile = ProcessInvoice;
                    base.ProcessSecondaryFile = null;

                    //Start processing all new data files dropped within the DataLocation folder.

                    FileWatcher1.EnableRaisingEvents = true;
                    FileWatcher2.EnableRaisingEvents = false;
                }
                catch (Exception ex)
                {
                    LogManager.WriteLog(DefaultErrorLogName, LogLocation, ex.Message, isError: true, stackTrace: ex.StackTrace, useDatedLogs: true, mirrorToConsole: LogToConsole);
                }
            }

            return ((stopProcessing != true) && (success == true));
        }

        /// <summary>
        /// Processes a file either found during start-up or dropped while monitoring
        /// the invoice drop location
        /// </summary>
        /// <param name="name">File Name</param>
        /// <param name="filePath">Full path to the invoice file</param>
        public void ProcessInvoice(string name, string filePath)
        {
            try
            {
                if ((String.IsNullOrEmpty(SubstitutionLocation) == true) ||
                    ((String.IsNullOrEmpty(SubstitutionLocation) != true) && (SubstitutionLocation.ToLower().Equals(filePath.ToLower()) != true)))
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

                        //If a substitution sheet exists, parse it so that its data can be used during the file
                        //parsing routine. This process is done each time a new input file is processed to allow the
                        //cheat sheet to be periodically updated without restarting the application.

                        List<Customer3InvoiceLineItemSubstitution> substitutions = new List<Customer3InvoiceLineItemSubstitution>();
                        if (String.IsNullOrEmpty(SubstitutionLocation) != true)
                        {
                            substitutions = ParseExcelSubstitutions();
                        }

                        //Parsing data from the physical file prior to storing it in the
                        //database.

                        List<Customer3Invoice> invoices = ParseExcelInvoice(database, filePath);

                        //Process all invoice substitutions based on the data provided by the
                        //Excel substitution sheet.

                        ProcessInvoiceSubstitutions(database, invoices, substitutions);

                        //Save all invoice information to the order / order detail tables in
                        //the Customer3 database.

                        Customer3OrderManager orderManager = new Customer3OrderManager(database);

                        foreach (Customer3Invoice invoice in invoices)
                        {
                            try
                            {
                                //Searching each line item for discrepencies between the stated part cost on the 
                                //PO and the Customer3 part cost in the database before saving.

                                foreach (Customer3InvoiceLineItem lineItem in invoice.LineItems)
                                {
                                    if ((lineItem.UnitCost.Equals(lineItem.Customer3UnitCost) != true) &&
                                        ((invoice.SalesOrderStatus.Equals("CANCELLED") != true) || (lineItem.LineNumberStatus.Equals("CANCELLED") != true)))
                                    {
                                        LogManager.WriteLog(_pricingLogName, LogLocation, String.Format("Part number {0} has a PO unit cost of {1} on purchase order {2} and a Customer3 unit cost of {3}.",
                                            lineItem.ItemNumber, lineItem.UnitCost, invoice.PurchaseOrder, lineItem.Customer3UnitCost), stackTrace: String.Empty);
                                    }
                                }

                                orderManager.Save(invoice);

                                //Add all appropriate log entries to document that this file was processed prior to
                                //moving it to the processed files location.

                                LogManager.WriteLog(DefaultErrorLogName, LogLocation, String.Format("Processing completed for purchase order '{0}' with {1} line items as part of file {2}", invoice.PurchaseOrder, invoice.LineItems.Count, name), mirrorToConsole: LogToConsole);
                            }
                            catch (Exception ex)
                            {
                                if (LogToConsole == true)
                                {
                                    Console.ForegroundColor = ConsoleColor.Red;
                                }

                                //Add all appropriate log entries to document that an error was encountered while processing
                                //invoice data within this file.

                                LogManager.WriteLog(DefaultErrorLogName, LogLocation, String.Format("An error was encountered while attempting to save purchase order '{0}' as part of file {1}; the error is on the next line.", invoice.PurchaseOrder, name), isError: true, mirrorToConsole: LogToConsole);
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
            }
            catch (Exception ex)
            {
                LogProcessingError(ex, name, filePath);
            }
        }

        #endregion

        #region Private Methods

        private void ProcessInvoiceSubstitutions(Database database, List<Customer3Invoice> invoices, List<Customer3InvoiceLineItemSubstitution> substitutions)
        {
            foreach (Customer3Invoice invoice in invoices)
            {
                List<Customer3InvoiceLineItem> lineItemSubstitutions = new List<Customer3InvoiceLineItem>();
                if (invoice.SalesOrderStatus.ToLower().Equals("open") == true)
                {
                    Customer3PartManager partManager = new Customer3PartManager(database);

                    foreach (Customer3InvoiceLineItem lineItem in invoice.LineItems)
                    {
                        if (lineItem.LineNumberStatus.ToLower().Equals("open") == true)
                        {
                            var substitution = substitutions.Where(x => x.ShippablePackItemNumber.Equals(lineItem.ItemNumber)).FirstOrDefault();

                            //If substitution values exist for this line item, process both the Import Kit substitutions
                            //as well as the Crossbar substitutions; leaving the original line item intact with a quantity of zero.

                            if (substitution != null)
                            {
                                int Customer3PartNumberID = partManager.GetPartIdentifier(substitution.ImportKitItemNumber);
                                decimal Customer3PartNumberUnitCost = partManager.GetPartUnitCost(substitution.ImportKitItemNumber);

                                //If the part number and unit cost are valid, continue with the substitution

                                if ((Customer3PartNumberID > 0) && (Customer3PartNumberUnitCost > 0))
                                {
                                    //Import Kit Substitution
                                    Customer3InvoiceLineItem importKitLineItem = new Customer3InvoiceLineItem();

                                    importKitLineItem.ItemNumber = substitution.ImportKitItemNumber;
                                    importKitLineItem.OrderQuantity = (substitution.ImportKitItemNumberQuantity * lineItem.OrderQuantity);
                                    importKitLineItem.Customer3PartNumberID = Customer3PartNumberID;
                                    importKitLineItem.Customer3UnitCost = Customer3PartNumberUnitCost;
                                    importKitLineItem.UnitCost = Customer3PartNumberUnitCost;
                                    importKitLineItem.LineNumber = lineItem.LineNumber;
                                    importKitLineItem.LineNumberStatus = lineItem.LineNumberStatus;
                                   
                                    lineItemSubstitutions.Add(importKitLineItem);
                                    LogManager.WriteLog(_pricingLogName, LogLocation, String.Format("Processing substitution for line item {0}; adding import kit line item {1} to purchase order {2} at line number {3}.",
                                        lineItem.ItemNumber, substitution.ImportKitItemNumber, invoice.PurchaseOrder, lineItem.LineNumber), mirrorToConsole: LogToConsole);
                                }

                                Customer3PartNumberID = partManager.GetPartIdentifier(substitution.CrossbarsItemNumber);
                                Customer3PartNumberUnitCost = partManager.GetPartUnitCost(substitution.CrossbarsItemNumber);

                                if ((Customer3PartNumberID > 0) && (Customer3PartNumberUnitCost > 0))
                                {
                                    //Crossbar Substitution
                                    Customer3InvoiceLineItem crossbarSubstitution = new Customer3InvoiceLineItem();

                                    crossbarSubstitution.ItemNumber = substitution.CrossbarsItemNumber;
                                    crossbarSubstitution.OrderQuantity = (substitution.CrossbarsItemNumberQuantity * lineItem.OrderQuantity);
                                    crossbarSubstitution.Customer3PartNumberID = Customer3PartNumberID;
                                    crossbarSubstitution.Customer3UnitCost = Customer3PartNumberUnitCost;
                                    crossbarSubstitution.UnitCost = Customer3PartNumberUnitCost;
                                    crossbarSubstitution.LineNumber = lineItem.LineNumber;
                                    crossbarSubstitution.LineNumberStatus = lineItem.LineNumberStatus;
                                   
                                    lineItemSubstitutions.Add(crossbarSubstitution);
                                    LogManager.WriteLog(_pricingLogName, LogLocation, String.Format("Processing substitution for line item {0}; adding crossbar line item {1} to purchase order {2} at line number {3}.",
                                        lineItem.ItemNumber, substitution.CrossbarsItemNumber, invoice.PurchaseOrder, lineItem.LineNumber), mirrorToConsole: LogToConsole);
                                }

                                //Zero out the quantity on the original line item in favor
                                //of the substitutions

                                lineItem.OrderQuantity = 0;
                            }
                        }
                    }

                    //Adding all substitutions in to the invoices line item list, if any
                    //were processed.

                    invoice.LineItems.AddRange(lineItemSubstitutions);
                }
            }
        }

        /// <summary>
        /// Parses the excel cheat sheet provided and returns a list of
        /// data elements for input file processing.
        /// </summary>
        private List<Customer3InvoiceLineItemSubstitution> ParseExcelSubstitutions()
        {
            List<Customer3InvoiceLineItemSubstitution> retVals = new List<Customer3InvoiceLineItemSubstitution>();
            string ext = Path.GetExtension(SubstitutionLocation);

            IWorkbook book = null;
            ISheet sheet = null;

            try
            {
                using (FileStream fs = new FileStream(SubstitutionLocation, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
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

                    int row = 0;
                    while (enumerator.MoveNext() == true)
                    {
                        row++;
                        IRow lineItem = (IRow)enumerator.Current;

                        if (lineItem == null)
                        {
                            throw new Exception("Unexpected null row data found in Excel file provided.");
                        }

                        if (row >= DataOrdinals.Substitution_Excel_StartsOnRow)
                        {
                            Customer3InvoiceLineItemSubstitution cheat = new Customer3InvoiceLineItemSubstitution();

                            //If the first cell to be parsed is blank, the entire line is blank; skip it.

                            if ((lineItem.Cells[DataOrdinals.Substitution_Excel_ShippablePackItemNumber].CellType == CellType.Blank) != true)
                            {
                                cheat.ShippablePackItemNumber = lineItem.Cells[DataOrdinals.Substitution_Excel_ShippablePackItemNumber].StringCellValue;
                                cheat.ShippablePackItemNumberDescription = lineItem.Cells[DataOrdinals.Substitution_Excel_ShippablePackItemNumberDescription].StringCellValue;

                                int index = cheat.ShippablePackItemNumber.IndexOf('*');
                                cheat.ShippablePackItemNumber = cheat.ShippablePackItemNumber.Substring(0, index);

                                cheat.ImportKitItemNumber = lineItem.Cells[DataOrdinals.Substitution_Excel_ImportKitItemNumber].StringCellValue;
                                cheat.ImportKitItemNumberDescription = lineItem.Cells[DataOrdinals.Substitution_Excel_ImportKitItemNumberDescription].StringCellValue;

                                index = cheat.ImportKitItemNumber.IndexOf('*');
                                cheat.ImportKitItemNumber = cheat.ImportKitItemNumber.Substring(0, index);

                                cheat.CrossbarsItemNumber = lineItem.Cells[DataOrdinals.Substitution_Excel_CrossbarsItemNumber].StringCellValue;
                                cheat.CrossbarsItemNumberDescription = lineItem.Cells[DataOrdinals.Substitution_Excel_CrossbarsItemNumberDescription].StringCellValue;

                                index = cheat.CrossbarsItemNumber.IndexOf('*');
                                cheat.CrossbarsItemNumber = cheat.CrossbarsItemNumber.Substring(0, index);

                                int importKitQuantity = 0;
                                int crossbarQuantity = 0;

                                if (lineItem.Cells[DataOrdinals.Substitution_Excel_ImportKitItemNumberQuantity].CellType == CellType.Numeric)
                                {
                                    double value = lineItem.Cells[DataOrdinals.Substitution_Excel_ImportKitItemNumberQuantity].NumericCellValue;
                                    importKitQuantity = value.ToString().ToInt32();
                                }
                                else
                                {
                                    string value = lineItem.Cells[DataOrdinals.Substitution_Excel_ImportKitItemNumberQuantity].StringCellValue;
                                    Int32.TryParse(value, out importKitQuantity);
                                }

                                if (lineItem.Cells[DataOrdinals.Substitution_Excel_CrossbarsItemNumberQuantity].CellType == CellType.Numeric)
                                {
                                    double value = lineItem.Cells[DataOrdinals.Substitution_Excel_CrossbarsItemNumberQuantity].NumericCellValue;
                                    crossbarQuantity = value.ToString().ToInt32();
                                }
                                else
                                {
                                    string value = lineItem.Cells[DataOrdinals.Substitution_Excel_CrossbarsItemNumberQuantity].StringCellValue;
                                    Int32.TryParse(value, out crossbarQuantity);
                                }

                                cheat.ImportKitItemNumberQuantity = importKitQuantity;
                                cheat.CrossbarsItemNumberQuantity = crossbarQuantity;

                                retVals.Add(cheat);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteLog(DefaultErrorLogName, LogLocation, String.Format("An error occured while attempting to process the configured cheat sheet, '{0}'; the error message follows on the next line.", SubstitutionLocation), isError: true, mirrorToConsole: LogToConsole);
                LogManager.WriteLog(DefaultErrorLogName, LogLocation, ex.Message, isError: true, stackTrace: ex.StackTrace, mirrorToConsole: LogToConsole);
            }
            finally
            {
                if (book != null)
                {
                    book.Close();
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return retVals;
        }

        /// <summary>
        /// Parses the excel invoice file provided using the current database connection for any 
        /// supplementary look-ups required.
        /// </summary>
        private List<Customer3Invoice> ParseExcelInvoice(Database database, string invoiceFilePath)
        {
            List<Customer3Invoice> retVal = new List<Customer3Invoice>();
            Exception caught = null;

            string ext = Path.GetExtension(invoiceFilePath);

            IWorkbook book = null;
            ISheet sheet = null;

            try
            {
                using (FileStream fs = new FileStream(invoiceFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
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

                        if (lineItem.Cells.Count < DataOrdinals.Invoice_Excel_NeededColumns)
                        {
                            throw new Exception("The number of columns expected in the Excel spreadsheet does not equal the number found.");
                        }

                        //Check if this is the header row; if so, by-pass it and begin processing
                        //the actual sheet data in subsequent rows.

                        if (lineItem.Cells[DataOrdinals.Invoice_Excel_Ordinal_Supplier].StringCellValue.ToLower().Equals("supplier name") != true)
                        {
                            int po = lineItem.Cells[DataOrdinals.Invoice_Excel_Ordinal_PurchaseOrder].StringCellValue.ToInt32();
                            Customer3Invoice invoice = retVal.Find(x => x.PurchaseOrder.Equals(po));

                            //If this purchase order is not yet within our list of invoices then this is the first time we've encountered
                            //it in the excel file; create a new record, import its values and process all of its line items.

                            if (invoice == null)
                            {
                                invoice = new Customer3Invoice();

                                invoice.PurchaseOrder = po;
                                invoice.SalesOrderStatus = lineItem.Cells[DataOrdinals.Invoice_Excel_Ordinal_HeaderStatus].StringCellValue.ToUpper();
                                invoice.SalesOrderNumber = lineItem.Cells[DataOrdinals.Invoice_Excel_Ordinal_OrderNumber].StringCellValue;
                                invoice.Revision = lineItem.Cells[DataOrdinals.Invoice_Excel_Ordinal_Revision].StringCellValue.ToInt32();
                                invoice.NeedByDate = lineItem.Cells[DataOrdinals.Invoice_Excel_Ordinal_NeedByDate].StringCellValue;
                                invoice.RouteCode = lineItem.Cells[DataOrdinals.Invoice_Excel_Ordinal_RouteCode].StringCellValue;

                                //This first record will also contain data for the first line item in the invoice; parse it out
                                //and add the line item record to the invoice.

                                Customer3InvoiceLineItem invoicelineItem = ParseExcelInvoiceLineItem(database, invoice, lineItem);
                                
                                //Setting the total cost based on this initial line item and adding the corresponding line item
                                //to the invoices internal collection.

                                invoice.TotalCost = (invoicelineItem.OrderQuantity * invoicelineItem.UnitCost);
                                invoice.TotalCustomer3Cost = (invoicelineItem.OrderQuantity * invoicelineItem.Customer3UnitCost);

                                invoice.LineItems.Add(invoicelineItem);

                                //Add this new invoice to the results collection

                                retVal.Add(invoice);
                            }
                            else
                            {
                                //If this purchase order already exists within our list of invoices then this is another line item on that
                                //invoice; create a new line item record, process its values and add it to the invoice.

                                Customer3InvoiceLineItem invoicelineItem = ParseExcelInvoiceLineItem(database, invoice, lineItem);

                                //Updating the total cost for this invoice based on this values in this line item and adding
                                //the line item to the invoices internal collection.

                                invoice.TotalCost += (invoicelineItem.OrderQuantity * invoicelineItem.UnitCost);
                                invoice.TotalCustomer3Cost += (invoicelineItem.OrderQuantity * invoicelineItem.Customer3UnitCost);

                                invoice.LineItems.Add(invoicelineItem);
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
        private Customer3InvoiceLineItem ParseExcelInvoiceLineItem(Database database, Customer3Invoice invoice, IRow row)
        {
            Customer3InvoiceLineItem retVal = new Customer3InvoiceLineItem();

            retVal.ItemNumber = row.Cells[DataOrdinals.Invoice_Excel_Ordinal_ItemNumber].StringCellValue;
            retVal.LineNumber = row.Cells[DataOrdinals.Invoice_Excel_Ordinal_LineNumber].StringCellValue;
            retVal.LineNumberStatus = row.Cells[DataOrdinals.Invoice_Excel_Ordinal_ItemStatus].StringCellValue.ToUpper();
            retVal.UnitCost = row.Cells[DataOrdinals.Invoice_Excel_Ordinal_UnitPrice].StringCellValue.ToDecimal();
            retVal.OrderQuantity = row.Cells[DataOrdinals.Invoice_Excel_Ordinal_Quantity].StringCellValue.ToInt32();

            //Parsing out unneeded characters in the item number

            int index = retVal.ItemNumber.IndexOf('*');
            retVal.ItemNumber = retVal.ItemNumber.Substring(0, index);

            //If this invoice or line number is in a cancelled statue, zero out the quantity and unit costs on this line item.

            if ((invoice.SalesOrderStatus.Equals("CANCELLED") == true) || (retVal.LineNumberStatus.Equals("CANCELLED") == true))
            {
                retVal.OrderQuantity = 0;
                retVal.Customer3UnitCost = 0;
                retVal.UnitCost = 0;
            }
            else
            {
                Customer3PartManager partManager = new Customer3PartManager(database);

                retVal.Customer3UnitCost = partManager.GetPartUnitCost(retVal.ItemNumber);
                retVal.Customer3PartNumberID = partManager.GetPartIdentifier(retVal.ItemNumber);
            }

            return retVal;
        }

        #endregion

        public static Customer3Processor CreateInstance()
        {
            return new Customer3Processor();
        }
    }
}
