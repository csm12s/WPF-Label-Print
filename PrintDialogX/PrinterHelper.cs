using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Printing;
using System.Reflection;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;
using System.Windows.Markup;
using System.Windows.Controls;
using System.Windows.Documents;
using Size = System.Windows.Size;
using MessageBox = System.Windows.MessageBox;
using Brushes = System.Windows.Media.Brushes;
using Point = System.Windows.Point;

namespace PrintDialogX.PrinterHelper
{
    public class PrinterHelper
    {
        /// <summary>
        /// Initialize a <see cref="PrinterHelper"/> class.
        /// </summary>
        protected PrinterHelper()
        {
            return;
        }

        /// <summary>
        /// Get the printer by the printer name.
        /// </summary>
        /// <param name="printerName">The printer name.</param>
        /// <returns>The printer.</returns>
        public static PrintQueue GetPrinterByName(string printerName)
        {
            return new PrintServer().GetPrintQueue(printerName);
        }

        /// <summary>
        /// Get the printer by the printer name.
        /// </summary>
        /// <param name="printerName">The printer name.</param>
        /// <param name="printServer">The print server that used to get printer.</param>
        /// <returns>The printer.</returns>
        public static PrintQueue GetPrinterByName(string printerName, PrintServer printServer)
        {
            return printServer.GetPrintQueue(printerName);
        }

        /// <summary>
        /// Get local default printer.
        /// </summary>
        /// <returns>The printer.</returns>
        public static PrintQueue GetDefaultPrinter()
        {
            return LocalPrintServer.GetDefaultPrintQueue();
        }

        /// <summary>
        /// Get local printers.
        /// </summary>
        /// <returns>A collection of printers.</returns>
        public static PrintQueueCollection GetLocalPrinters()
        {
            return GetLocalPrinters(new PrintServer());
        }

        /// <summary>
        /// Get local printers.
        /// </summary>
        /// <param name="enumerationFlag">An array of values that represent the types of print queues that are in the collection.</param>
        /// <returns>A collection of printers.</returns>
        public static PrintQueueCollection GetLocalPrinters(EnumeratedPrintQueueTypes[] enumerationFlag)
        {
            return GetLocalPrinters(new PrintServer(), enumerationFlag);
        }

        /// <summary>
        /// Get local printers.
        /// </summary>
        /// <param name="server">The print server.</param>
        /// <returns>A collection of printers.</returns>
        public static PrintQueueCollection GetLocalPrinters(PrintServer server)
        {
            return server.GetPrintQueues();
        }

        /// <summary>
        /// Get local printers.
        /// </summary>
        /// <param name="server">The print server.</param>
        /// <param name="enumerationFlag">An array of values that represent the types of print queues that are in the collection.</param>
        /// <returns>A collection of printers.</returns>
        public static PrintQueueCollection GetLocalPrinters(PrintServer server, EnumeratedPrintQueueTypes[] enumerationFlag)
        {
            return server.GetPrintQueues(enumerationFlag);
        }

        /// <summary>
        /// Get printer's status info.
        /// </summary>
        /// <param name="printerName">Printer's name.</param>
        /// <returns>printer's status info.</returns>
        public static string GetPrinterStatusInfo(string printerName)
        {
            return GetPrinterStatusInfo(new PrintServer().GetPrintQueue(printerName));
        }

        /// <summary>
        /// Get printer's status info.
        /// </summary>
        /// <param name="printer">Printer.</param>
        /// <returns>printer's status info.</returns>
        public static string GetPrinterStatusInfo(PrintQueue printer)
        {
            printer.Refresh();

            return GetPrinterStatusInfo(printer.QueueStatus);
        }

        /// <summary>
        /// Get printer's status info.
        /// </summary>
        /// <param name="printerStatue">Printer's status.</param>
        /// <returns>printer's status info.</returns>
        public static string GetPrinterStatusInfo(PrintQueueStatus printerStatue)
        {
            return printerStatue switch
            {
                PrintQueueStatus.Busy => "Busy",
                PrintQueueStatus.DoorOpen => "Door Open",
                PrintQueueStatus.Error => "Error",
                PrintQueueStatus.Initializing => "Initializing",
                PrintQueueStatus.IOActive => "Exchanging Data",
                PrintQueueStatus.ManualFeed => "Need Manual Feed",
                PrintQueueStatus.NoToner => "No Toner",
                PrintQueueStatus.Offline => "Offline",
                PrintQueueStatus.OutOfMemory => "Out Of Memory",
                PrintQueueStatus.OutputBinFull => "Output Bin Full",
                PrintQueueStatus.PagePunt => "Page Punt",
                PrintQueueStatus.PaperJam => "Paper Jam",
                PrintQueueStatus.PaperOut => "Paper Out",
                PrintQueueStatus.PaperProblem => "Paper Error",
                PrintQueueStatus.Paused => "Paused",
                PrintQueueStatus.PendingDeletion => "Deleting Job",
                PrintQueueStatus.PowerSave => "Power Save",
                PrintQueueStatus.Printing => "Printing",
                PrintQueueStatus.Processing => "Processing",
                PrintQueueStatus.ServerUnknown => "Server Unknown",
                PrintQueueStatus.TonerLow => "Toner Low",
                PrintQueueStatus.UserIntervention => "Need User Intervention",
                PrintQueueStatus.Waiting => "Waiting",
                PrintQueueStatus.WarmingUp => "Warming Up",

                _ => "Ready",
            };
        }

        /// <summary>
        /// Install new printer to the print server.
        /// </summary>
        /// <param name="printerName">The new printer name.</param>
        /// <param name="driverName">The new printer driver name.</param>
        /// <param name="portNames">IDs of the ports that the new queue uses.</param>
        /// <param name="printerProcessorName">The print processor name.</param>
        /// <param name="printerProperties">The new printer properties.</param>
        public static void InstallPrinter(string printerName, string driverName, string[] portNames, string printerProcessorName, PrintQueueAttributes printerProperties)
        {
            LocalPrintServer localPrintServer = new LocalPrintServer();
            localPrintServer.InstallPrintQueue(printerName, driverName, portNames, printerProcessorName, printerProperties);
            localPrintServer.Commit();
        }

        /// <summary>
        /// Install new printer to the print server.
        /// </summary>
        /// <param name="printServer">The print server.</param>
        /// <param name="printerName">The new printer name.</param>
        /// <param name="driverName">The new printer driver name.</param>
        /// <param name="portNames">IDs of the ports that the new queue uses.</param>
        /// <param name="printerProcessorName">The print processor name.</param>
        /// <param name="printerProperties">The new printer properties.</param>
        public static void InstallPrinter(PrintServer printServer, string printerName, string driverName, string[] portNames, string printerProcessorName, PrintQueueAttributes printerProperties)
        {
            printServer.InstallPrintQueue(printerName, driverName, portNames, printerProcessorName, printerProperties);
            printServer.Commit();
        }

        /// <summary>
        /// Install new printer to the print server.
        /// </summary>
        /// <param name="printServer">The print server.</param>
        /// <param name="printerName">The new printer name.</param>
        /// <param name="driverName">The new printer driver name.</param>
        /// <param name="portNames">IDs of the ports that the new queue uses.</param>
        /// <param name="printerProcessorName">The print processor name.</param>
        /// <param name="printerProperties">The new printer properties.</param>
        /// <param name="printerShareName">The new printer share name.</param>
        /// <param name="printerComment">The new printer comment.</param>
        /// <param name="printerLoction">The new printer loction.</param>
        /// <param name="printerSeparatorFile">The path of a file that is inserted at the beginning of each print job.</param>
        /// <param name="printerPriority">A value from 1 through 99 that specifies the priority of the queue relative to other queues that are hosted by the print server.</param>
        /// <param name="printerDefaultPriority">A value from 1 through 99 that specifies the default priority of new print jobs that are sent to the queue.</param>
        public static void InstallPrinter(PrintServer printServer, string printerName, string driverName, string[] portNames, string printerProcessorName, PrintQueueAttributes printerProperties, string printerShareName, string printerComment, string printerLoction, string printerSeparatorFile, int printerPriority, int printerDefaultPriority)
        {
            printServer.InstallPrintQueue(printerName, driverName, portNames, printerProcessorName, printerProperties, printerShareName, printerComment, printerLoction, printerSeparatorFile, printerPriority, printerDefaultPriority);
            printServer.Commit();
        }
    }

    public class PrintJobHelper
    {
        /// <summary>
        /// Initialize a <see cref="PrintJobHelper"/> class.
        /// </summary>
        protected PrintJobHelper()
        {
            return;
        }

        /// <summary>
        /// Get all print jobs of printer.
        /// </summary>
        /// <param name="printerName">The printer name.</param>
        /// <returns>The print jobs.</returns>
        public static PrintJobInfoCollection GetPrintJobs(string printerName)
        {
            PrintQueue printer = new PrintServer().GetPrintQueue(printerName);
            printer.Refresh();

            return printer.GetPrintJobInfoCollection();
        }

        /// <summary>
        /// Get all print jobs of printer.
        /// </summary>
        /// <param name="printer">The printer.</param>
        /// <returns>The print jobs.</returns>
        public static PrintJobInfoCollection GetPrintJobs(PrintQueue printer)
        {
            printer.Refresh();

            return printer.GetPrintJobInfoCollection();
        }

        /// <summary>
        /// Get all print jobs of printer.
        /// </summary>
        /// <param name="printer">The printer.</param>
        /// <param name="submitter">The print job submitter.</param>
        /// <returns>The print jobs.</returns>
        public static PrintJobInfoCollection GetPrintJobs(PrintQueue printer, string submitter)
        {
            printer.Refresh();

            PrintJobInfoCollection printJobList = new PrintJobInfoCollection(printer, null);

            foreach (PrintSystemJobInfo jobInfo in printer.GetPrintJobInfoCollection())
            {
                if (jobInfo.Submitter == submitter)
                {
                    printJobList.Add(jobInfo);
                }
            }

            return printJobList;
        }

        /// <summary>
        /// Get all error print jobs of printer.
        /// </summary>
        /// <param name="printerName">The printer name.</param>
        /// <returns>The print jobs.</returns>
        public static PrintJobInfoCollection GetErrorPrintJobs(string printerName)
        {
            return GetErrorPrintJobs(new PrintServer().GetPrintQueue(printerName), false);
        }

        /// <summary>
        /// Get all error print jobs of printer.
        /// </summary>
        /// <param name="printer">The printer.</param>
        /// <returns>The print jobs.</returns>
        public static PrintJobInfoCollection GetErrorPrintJobs(PrintQueue printer)
        {
            return GetErrorPrintJobs(printer, false);
        }

        /// <summary>
        /// Get all error print jobs of printer.
        /// </summary>
        /// <param name="printer">The printer.</param>
        /// <param name="isCancel">Cancel error print jobs or not.</param>
        /// <returns>The print jobs.</returns>
        public static PrintJobInfoCollection GetErrorPrintJobs(PrintQueue printer, bool isCancel)
        {
            printer.Refresh();

            PrintJobInfoCollection errorList = new PrintJobInfoCollection(printer, null);

            foreach (PrintSystemJobInfo jobInfo in printer.GetPrintJobInfoCollection())
            {
                if (jobInfo.JobStatus == PrintJobStatus.Blocked || jobInfo.JobStatus == PrintJobStatus.Error || jobInfo.JobStatus == PrintJobStatus.Offline || jobInfo.JobStatus == PrintJobStatus.PaperOut || jobInfo.JobStatus == PrintJobStatus.UserIntervention)
                {
                    errorList.Add(jobInfo);

                    if (isCancel)
                    {
                        jobInfo.Cancel();
                        jobInfo.Commit();
                    }
                }
            }

            return errorList;
        }

        /// <summary>
        /// Get all error print jobs of printer.
        /// </summary>
        /// <param name="printer">The printer.</param>
        /// <param name="submitter">The print job submitter.</param>
        /// <returns>The print jobs.</returns>
        public static PrintJobInfoCollection GetErrorPrintJobs(PrintQueue printer, string submitter)
        {
            return GetErrorPrintJobs(printer, submitter, false);
        }

        /// <summary>
        /// Get all error print jobs of printer.
        /// </summary>
        /// <param name="printer">The printer.</param>
        /// <param name="submitter">The print job submitter.</param>
        /// <param name="isCancel">Cancel error print jobs or not.</param>
        /// <returns>The print jobs.</returns>
        public static PrintJobInfoCollection GetErrorPrintJobs(PrintQueue printer, string submitter, bool isCancel)
        {
            printer.Refresh();

            PrintJobInfoCollection errorList = new PrintJobInfoCollection(printer, null);

            foreach (PrintSystemJobInfo jobInfo in printer.GetPrintJobInfoCollection())
            {
                if (jobInfo.Submitter == submitter)
                {
                    if (jobInfo.JobStatus == PrintJobStatus.Blocked || jobInfo.JobStatus == PrintJobStatus.Error || jobInfo.JobStatus == PrintJobStatus.Offline || jobInfo.JobStatus == PrintJobStatus.PaperOut || jobInfo.JobStatus == PrintJobStatus.UserIntervention)
                    {
                        errorList.Add(jobInfo);

                        if (isCancel)
                        {
                            jobInfo.Cancel();
                            jobInfo.Commit();
                        }
                    }
                }
            }

            return errorList;
        }
    }
}

namespace PrintDialogX.PaperHelper
{
    public class PaperHelper
    {
        /// <summary>
        /// Initialize a <see cref="PaperHelper"/> class.
        /// </summary>
        protected PaperHelper()
        {
            return;
        }

        /// <summary>
        /// Get actual paper size.
        /// </summary>
        /// <param name="pageSizeName">Paper size name.</param>
        /// <returns>Paper size.</returns>
        [Obsolete]
        public static Size GetPaperSize(PageMediaSizeName pageSizeName)
        {
            System.Windows.Controls.PrintDialog printDlg = new System.Windows.Controls.PrintDialog();
            PageMediaSize pageMediaSize = new PageMediaSize(pageSizeName);
            printDlg.PrintTicket.PageMediaSize = pageMediaSize;

            return new Size(printDlg.PrintableAreaWidth, printDlg.PrintableAreaHeight);
        }

        /// <summary>
        /// Get actual paper size.
        /// </summary>
        /// <param name="pageSizeName">Paper size name.</param>
        /// <param name="pageOrientation">Paper orientation.</param>
        /// <returns>Paper size.</returns>
        [Obsolete]
        public static Size GetPaperSize(PageMediaSizeName pageSizeName, PrintSettings.PageOrientation pageOrientation)
        {
            System.Windows.Controls.PrintDialog printDlg = new System.Windows.Controls.PrintDialog();
            PageMediaSize pageMediaSize = new PageMediaSize(pageSizeName);
            printDlg.PrintTicket.PageMediaSize = pageMediaSize;

            if (pageOrientation == PrintSettings.PageOrientation.Portrait)
            {
                return new Size(printDlg.PrintableAreaWidth, printDlg.PrintableAreaHeight);
            }
            else
            {
                return new Size(printDlg.PrintableAreaHeight, printDlg.PrintableAreaWidth);
            }
        }

        /// <summary>
        /// Get actual paper size.
        /// </summary>
        /// <param name="CMWidth">Paper width in cm.</param>
        /// <param name="CMHeight">Paper height in cm.</param>
        /// <returns>Paper size.</returns>
        public static Size GetPaperSize(double CMWidth, double CMHeight)
        {
            double cm = 37.7952755905512;
            return new Size(CMWidth * cm / 10, CMHeight * cm / 10);
        }

        /// <summary>
        /// Get actual paper size.
        /// </summary>
        /// <param name="pageSizeName">Paper size name.</param>
        /// <param name="isAdvanced">Use advanced calculate formula or not.</param>
        /// <returns>Paper size.</returns>
        public static Size GetPaperSize(PageMediaSizeName pageSizeName, bool isAdvanced)
        {
            System.Windows.Controls.PrintDialog printDlg = new System.Windows.Controls.PrintDialog();
            PageMediaSize pageMediaSize = new PageMediaSize(pageSizeName);
            printDlg.PrintTicket.PageMediaSize = pageMediaSize;

            if (isAdvanced)
            {
                List<Size> printableAreaSize = new List<Size>();

                foreach (PrintQueue printer in PrinterHelper.PrinterHelper.GetLocalPrinters())
                {
                    printDlg.PrintQueue = printer;
                    printableAreaSize.Add(new Size((int)printDlg.PrintableAreaWidth, (int)printDlg.PrintableAreaHeight));
                }

                return printableAreaSize.GroupBy(i => i).OrderByDescending(group => group.Count()).Select(group => group.Key).First();
            }
            else
            {
                return new Size(printDlg.PrintableAreaHeight, printDlg.PrintableAreaWidth);
            }
        }

        /// <summary>
        /// Get actual paper size.
        /// </summary>
        /// <param name="pageSizeName">Paper size name.</param>
        /// <param name="printQueue">Printer that the size get from.</param>
        /// <returns>Paper size.</returns>
        public static Size GetPaperSize(PageMediaSizeName pageSizeName, PrintQueue printQueue)
        {
            System.Windows.Controls.PrintDialog printDlg = new System.Windows.Controls.PrintDialog();
            PageMediaSize pageMediaSize = new PageMediaSize(pageSizeName);
            printDlg.PrintQueue = printQueue;
            printDlg.PrintTicket.PageMediaSize = pageMediaSize;

            return new Size(printDlg.PrintableAreaWidth, printDlg.PrintableAreaHeight);
        }

        /// <summary>
        /// Get actual paper size.
        /// </summary>
        /// <param name="pageSizeName">Paper size name.</param>
        /// <param name="pageOrientation">Paper orientation.</param>
        /// <param name="isAdvanced">Use advanced calculate formula pr not.</param>
        /// <returns>Paper size.</returns>
        public static Size GetPaperSize(PageMediaSizeName pageSizeName, PrintSettings.PageOrientation pageOrientation, bool isAdvanced)
        {
            System.Windows.Controls.PrintDialog printDlg = new System.Windows.Controls.PrintDialog();
            PageMediaSize pageMediaSize = new PageMediaSize(pageSizeName);
            printDlg.PrintTicket.PageMediaSize = pageMediaSize;

            if (isAdvanced)
            {
                List<Size> printableAreaSize = new List<Size>();

                foreach (PrintQueue printer in PrinterHelper.PrinterHelper.GetLocalPrinters())
                {
                    printDlg.PrintQueue = printer;
                    printableAreaSize.Add(new Size((int)printDlg.PrintableAreaWidth, (int)printDlg.PrintableAreaHeight));
                }

                Size size = printableAreaSize.GroupBy(i => i).OrderByDescending(group => group.Count()).Select(group => group.Key).First();

                if (pageOrientation == PrintSettings.PageOrientation.Portrait)
                {
                    return size;
                }
                else
                {
                    return new Size(size.Height, size.Width);
                }
            }
            else
            {
                if (pageOrientation == PrintSettings.PageOrientation.Portrait)
                {
                    return new Size(printDlg.PrintableAreaWidth, printDlg.PrintableAreaHeight);
                }
                else
                {
                    return new Size(printDlg.PrintableAreaHeight, printDlg.PrintableAreaWidth);
                }
            }
        }

        /// <summary>
        /// Get actual paper size.
        /// </summary>
        /// <param name="pageSizeName">Paper size name.</param>
        /// <param name="pageOrientation">Paper orientation.</param>
        /// <param name="printQueue">Printer that the size get from.</param>
        /// <returns>Paper size.</returns>
        public static Size GetPaperSize(PageMediaSizeName pageSizeName, PrintSettings.PageOrientation pageOrientation, PrintQueue printQueue)
        {
            System.Windows.Controls.PrintDialog printDlg = new System.Windows.Controls.PrintDialog();
            PageMediaSize pageMediaSize = new PageMediaSize(pageSizeName);
            printDlg.PrintQueue = printQueue;
            printDlg.PrintTicket.PageMediaSize = pageMediaSize;

            if (pageOrientation == PrintSettings.PageOrientation.Portrait)
            {
                return new Size(printDlg.PrintableAreaWidth, printDlg.PrintableAreaHeight);
            }
            else
            {
                return new Size(printDlg.PrintableAreaHeight, printDlg.PrintableAreaWidth);
            }
        }
    }

    public class PaperSize
    {
        /// <summary>
        /// Initialize a <see cref="PaperSize"/> class.
        /// </summary>
        protected PaperSize()
        {
            return;
        }

        /// <summary>
        /// A1 paper width (mm).
        /// </summary>
        public static double A1Width { get; } = 594;

        /// <summary>
        /// A1 paper height (mm).
        /// </summary>
        public static double A1Height { get; } = 841;

        /// <summary>
        /// A2 paper width (mm).
        /// </summary>
        public static double A2Width { get; } = 420;

        /// <summary>
        /// A2 paper height (mm).
        /// </summary>
        public static double A2Height { get; } = 594;

        /// <summary>
        /// A3 paper width (mm).
        /// </summary>
        public static double A3Width { get; } = 297;

        /// <summary>
        /// A3 paper height (mm).
        /// </summary>
        public static double A3Height { get; } = 420;

        /// <summary>
        /// A4 paper width (mm).
        /// </summary>
        public static double A4Width { get; } = 210;

        /// <summary>
        /// A4 paper height (mm).
        /// </summary>
        public static double A4Height { get; } = 297;

        /// <summary>
        /// A5 paper width (mm).
        /// </summary>
        public static double A5Width { get; } = 148;

        /// <summary>
        /// A5 paper height (mm).
        /// </summary>
        public static double A5Height { get; } = 210;

        /// <summary>
        /// A6 paper width (mm).
        /// </summary>
        public static double A6Width { get; } = 105;

        /// <summary>
        /// A6 paper height (mm).
        /// </summary>
        public static double A6Height { get; } = 148;

        /// <summary>
        /// Search all width and height by name and return the result length. If can't find the length, it will return <see cref="Double.NaN"/>.
        /// </summary>
        /// <param name="name">The length name, allowed white-space and lowercase, like "A4 Height", "A4Height" and "a4 height" both means the A4 paper's height.</param>
        /// <returns>The length.</returns>
        public static double GetLengthWithName(string name)
        {
            foreach (PropertyInfo item in typeof(PaperSize).GetProperties(BindingFlags.Instance | BindingFlags.Static))
            {
                if (item.Name.ToLower() == name.Trim().ToLower())
                {
                    return (double)item.GetValue(new PaperSize());
                }
            }

            return double.NaN;
        }
    }
}

namespace PrintDialogX.DocumentMaker
{
    public class DocumentMaker
    {
        /// <summary>
        /// Initialize a <see cref="DocumentMaker"/> class.
        /// </summary>
        protected DocumentMaker()
        {
            return;
        }

        /// <summary>
        /// Make a document that contains an auto pagination <see cref="System.Windows.UIElement"/>.
        /// </summary>
        /// <param name="xaml">The XAML that represents the <see cref="System.Windows.UIElement"/>.</param>
        /// <param name="elementActualHeight">The <see cref="System.Windows.UIElement"/> actual height.</param>
        /// <returns>The document.</returns>
        public static FixedDocument PaginatonUIElementDocumentMaker(string xaml, double elementActualHeight)
        {
            return PaginatonUIElementDocumentMaker(xaml, elementActualHeight, null, 34, PageMediaSizeName.ISOA4, PrintSettings.PageOrientation.Portrait);
        }

        /// <summary>
        /// Make a document that contains an auto pagination <see cref="System.Windows.UIElement"/>.
        /// </summary>
        /// <param name="xaml">The XAML that represents the <see cref="System.Windows.UIElement"/>.</param>
        /// <param name="elementActualHeight">The <see cref="System.Windows.UIElement"/> actual height.</param>
        /// <returns>The document.</returns>
        public static FixedDocument PaginatonUIElementDocumentMaker(string xaml, double elementActualHeight, object dataContext)
        {
            return PaginatonUIElementDocumentMaker(xaml, elementActualHeight, dataContext, 34, PageMediaSizeName.ISOA4, PrintSettings.PageOrientation.Portrait);
        }

        /// <summary>
        /// Make a document that contains an auto pagination <see cref="System.Windows.UIElement"/>.
        /// </summary>
        /// <param name="xaml">The XAML that represents the <see cref="System.Windows.UIElement"/>.</param>
        /// <param name="elementActualHeight">The <see cref="System.Windows.UIElement"/> actual height.</param>
        /// <param name="dataContext">The <see cref="System.Windows.UIElement"/> data context.</param>
        /// <param name="documentMargin">The document margin.</param>
        /// <param name="documentSize">The document size.</param>
        /// <param name="documentOrientation">The document orientation.</param>
        /// <returns>The document.</returns>
        public static FixedDocument PaginatonUIElementDocumentMaker(string xaml, double elementActualHeight, object dataContext, double documentMargin, PageMediaSizeName documentSize, PrintSettings.PageOrientation documentOrientation)
        {
            return PaginatonUIElementDocumentMaker(xaml, elementActualHeight, dataContext, documentMargin, PaperHelper.PaperHelper.GetPaperSize(documentSize, true), documentOrientation);
        }

        /// <summary>
        /// Make a document that contains an auto pagination <see cref="System.Windows.UIElement"/>.
        /// </summary>
        /// <param name="xaml">The XAML that represents the <see cref="System.Windows.UIElement"/>.</param>
        /// <param name="elementActualHeight">The <see cref="System.Windows.UIElement"/> actual height.</param>
        /// <param name="dataContext">The <see cref="System.Windows.UIElement"/> data context.</param>
        /// <param name="documentMargin">The document margin.</param>
        /// <param name="documentSize">The document size.</param>
        /// <param name="documentOrientation">The document orientation.</param>
        /// <returns>The document.</returns>
        public static FixedDocument PaginatonUIElementDocumentMaker(string xaml, double elementActualHeight, object dataContext, double documentMargin, Size documentSize, PrintSettings.PageOrientation documentOrientation)
        {
            FixedDocument doc = new FixedDocument();
            if (documentOrientation == PrintSettings.PageOrientation.Portrait)
            {
                doc.DocumentPaginator.PageSize = documentSize;
            }
            else
            {
                doc.DocumentPaginator.PageSize = new Size(documentSize.Height, documentSize.Width);
            }

            int totalPage = (int)Math.Ceiling(elementActualHeight / (doc.DocumentPaginator.PageSize.Height - documentMargin * 2));

            for (int i = 0; i < totalPage; i++)
            {
                FrameworkElement printedElement = XamlReader.Parse(xaml) as FrameworkElement;
                printedElement.DataContext = dataContext;

                FixedPage fixedPage = new FixedPage
                {
                    Width = doc.DocumentPaginator.PageSize.Width,
                    Height = doc.DocumentPaginator.PageSize.Height
                };

                printedElement.Width = fixedPage.Width - documentMargin * 2;
                printedElement.Height = elementActualHeight;

                FixedPage.SetLeft(printedElement, documentMargin);
                FixedPage.SetTop(printedElement, -(i * (fixedPage.Height - documentMargin * 2)) + documentMargin);

                Grid clipGrid1 = new Grid()
                {
                    Width = fixedPage.Width,
                    Height = documentMargin,
                    Background = Brushes.White
                };

                Grid clipGrid2 = new Grid()
                {
                    Width = fixedPage.Width,
                    Height = documentMargin,
                    Background = Brushes.White
                };

                FixedPage.SetLeft(clipGrid1, 0);
                FixedPage.SetTop(clipGrid1, 0);

                FixedPage.SetLeft(clipGrid2, 0);
                FixedPage.SetTop(clipGrid2, fixedPage.Height - documentMargin);

                fixedPage.Children.Add(printedElement);
                fixedPage.Children.Add(clipGrid1);
                fixedPage.Children.Add(clipGrid2);

                fixedPage.Measure(doc.DocumentPaginator.PageSize);
                fixedPage.Arrange(new Rect(new Point(), doc.DocumentPaginator.PageSize));

                doc.Pages.Add(new PageContent() { Child = fixedPage });

                fixedPage.UpdateLayout();
                Internal.PrintPage.DoEvents();
            }

            return doc;
        }

        /// <summary>
        /// Make a document that contains an auto pagination <see cref="System.Windows.UIElement"/>.
        /// </summary>
        /// <param name="xaml">The XAML that represents the <see cref="System.Windows.UIElement"/>.</param>
        /// <param name="pageCount">The total page count.</param>
        /// <returns>The document.</returns>
        public static FixedDocument PaginatonUIElementDocumentMaker(string xaml, int pageCount)
        {
            return PaginatonUIElementDocumentMaker(xaml, pageCount, null, 34, PageMediaSizeName.ISOA4, PrintSettings.PageOrientation.Portrait);
        }

        /// <summary>
        /// Make a document that contains an auto pagination <see cref="System.Windows.UIElement"/>.
        /// </summary>
        /// <param name="xaml">The XAML that represents the <see cref="System.Windows.UIElement"/>.</param>
        /// <param name="pageCount">The total page count.</param>
        /// <returns>The document.</returns>
        public static FixedDocument PaginatonUIElementDocumentMaker(string xaml, int pageCount, object dataContext)
        {
            return PaginatonUIElementDocumentMaker(xaml, pageCount, dataContext, 34, PageMediaSizeName.ISOA4, PrintSettings.PageOrientation.Portrait);
        }

        /// <summary>
        /// Make a document that contains an auto pagination <see cref="System.Windows.UIElement"/>.
        /// </summary>
        /// <param name="xaml">The XAML that represents the <see cref="System.Windows.UIElement"/>.</param>
        /// <param name="pageCount">The total page count.</param>
        /// <param name="dataContext">The <see cref="System.Windows.UIElement"/> data context.</param>
        /// <param name="documentMargin">The document margin.</param>
        /// <param name="documentSize">The document size.</param>
        /// <param name="documentOrientation">The document orientation.</param>
        /// <returns>The document.</returns>
        public static FixedDocument PaginatonUIElementDocumentMaker(string xaml, int pageCount, object dataContext, double documentMargin, PageMediaSizeName documentSize, PrintSettings.PageOrientation documentOrientation)
        {
            FixedDocument doc = new FixedDocument();
            doc.DocumentPaginator.PageSize = PaperHelper.PaperHelper.GetPaperSize(documentSize, documentOrientation, true);

            int totalPage = pageCount;

            for (int i = 0; i < totalPage; i++)
            {
                FrameworkElement printedElement = XamlReader.Parse(xaml) as FrameworkElement;
                printedElement.DataContext = dataContext;

                FixedPage fixedPage = new FixedPage
                {
                    Width = doc.DocumentPaginator.PageSize.Width,
                    Height = doc.DocumentPaginator.PageSize.Height
                };

                printedElement.Width = fixedPage.Width - documentMargin * 2;
                printedElement.Height = double.NaN;

                FixedPage.SetLeft(printedElement, documentMargin);
                FixedPage.SetTop(printedElement, -(i * (fixedPage.Height - documentMargin * 2)) + documentMargin);

                Grid clipGrid1 = new Grid()
                {
                    Width = fixedPage.Width,
                    Height = documentMargin,
                    Background = Brushes.White
                };

                Grid clipGrid2 = new Grid()
                {
                    Width = fixedPage.Width,
                    Height = documentMargin,
                    Background = Brushes.White
                };

                FixedPage.SetLeft(clipGrid1, 0);
                FixedPage.SetTop(clipGrid1, 0);

                FixedPage.SetLeft(clipGrid2, 0);
                FixedPage.SetTop(clipGrid2, fixedPage.Height - documentMargin);

                fixedPage.Children.Add(printedElement);
                fixedPage.Children.Add(clipGrid1);
                fixedPage.Children.Add(clipGrid2);

                fixedPage.Measure(doc.DocumentPaginator.PageSize);
                fixedPage.Arrange(new Rect(new Point(), doc.DocumentPaginator.PageSize));

                doc.Pages.Add(new PageContent() { Child = fixedPage });

                fixedPage.UpdateLayout();
                Internal.PrintPage.DoEvents();
            }

            return doc;
        }
    }
}

namespace PrintDialogX.PrintDialogExceptions
{
    public class DocumentEmptyException : Exception
    {
        /// <summary>
        /// The message that describes the current exception.
        /// </summary>
        public override string Message { get; }

        /// <summary>
        /// The <see cref="FixedDocument"/> that is empty.
        /// </summary>
        public FixedDocument Document { get; }

        /// <summary>
        /// Initialize a <see cref="DocumentEmptyException"/> class.
        /// </summary>
        /// <param name="message">The message that describes the current exception.</param>
        /// <param name="document">The FixedDocument that is empty.</param>
        public DocumentEmptyException(string message, FixedDocument document)
        {
            Message = message;
            Document = document;
        }
    }

    public class UndefinedException : Exception
    {
        /// <summary>
        /// The message that describes the current exception.
        /// </summary>
        public override string Message { get; }

        /// <summary>
        /// The <see cref="System.Exception"/> instance that caused the current exception.
        /// </summary>
        public new Exception InnerException { get; }

        /// <summary>
        /// The time that exception throw out.
        /// </summary>
        public DateTime ExceptionTime { get; }

        /// <summary>
        /// Initialize a <see cref="UndefinedException"/> class.
        /// </summary>
        /// <param name="message">The message that describes the current exception.</param>
        public UndefinedException(string message)
        {
            Message = message;
            ExceptionTime = DateTime.Now;
        }

        /// <summary>
        /// Initialize a <see cref="UndefinedException"/> class.
        /// </summary>
        /// <param name="message">The message that describes the current exception.</param>
        /// <param name="innerException">The undefined exception.</param>
        public UndefinedException(string message, Exception innerException)
        {
            Message = message;
            InnerException = innerException;
            ExceptionTime = DateTime.Now;
        }
    }
}
