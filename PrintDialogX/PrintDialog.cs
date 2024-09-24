using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;
using System.Windows.Documents;
using Size = System.Windows.Size;
using MessageBox = System.Windows.MessageBox;
using Brushes = System.Windows.Media.Brushes;
using Point = System.Windows.Point;
using System.Reflection;

namespace PrintDialogX.PrintDialog
{
    public class PrintDialogSettings
    {
        /// <summary>
        /// Initialize a <see cref="PrintDialogSettings"/> class.
        /// </summary>
        public PrintDialogSettings()
        {//todo change
            UsePrinterDefaultSettings = false;
            Layout = PrintSettings.PageOrientation.Portrait;
            Color = PrintSettings.PageColor.Color;
            Quality = PrintSettings.PageQuality.Normal;
            NameID = 0;//page size name id
            //PageSize = PrintSettings.PageSize.ISOA4;
            PageType = PrintSettings.PageType.Plain;
            TwoSided = PrintSettings.TwoSided.OneSided;
            PagesPerSheet = 1;
            PageOrder = PrintSettings.PageOrder.Horizontal;
        }

        /// <summary>
        /// Use printer default settings or not.
        /// </summary>
        public bool UsePrinterDefaultSettings { get; internal set; } = false;

        /// <summary>
        /// Layout.
        /// </summary>
        public PrintSettings.PageOrientation Layout { get; set; } = PrintSettings.PageOrientation.Portrait;

        /// <summary>
        /// Color.
        /// </summary>
        public PrintSettings.PageColor Color { get; set; } = PrintSettings.PageColor.Color;

        /// <summary>
        /// Quality.
        /// </summary>
        public PrintSettings.PageQuality Quality { get; set; } = PrintSettings.PageQuality.Normal;

        /// <summary>
        /// Page size.
        /// </summary>

        public int NameID { get; set; } = 0;

        /// <summary>
        /// Page type.
        /// </summary>
        public PrintSettings.PageType PageType { get; set; } = PrintSettings.PageType.Plain;

        /// <summary>
        /// Two-sided.
        /// </summary>
        public PrintSettings.TwoSided TwoSided { get; set; } = PrintSettings.TwoSided.OneSided;

        /// <summary>
        /// Pages per sheet.
        /// </summary>
        public int PagesPerSheet { get; set; } = 1;

        /// <summary>
        /// Page order.
        /// </summary>
        public PrintSettings.PageOrder PageOrder { get; set; } = PrintSettings.PageOrder.Horizontal;

        /// <summary>
        /// Initialize a <see cref="PrintDialogSettings"/> that use printer default settings
        /// </summary>
        /// <returns>The <see cref="PrintDialogSettings"/> that use printer default settings.</returns>
        public static PrintDialogSettings PrinterDefaultSettings()
        {
            return new PrintDialogSettings()
            {
                UsePrinterDefaultSettings = true
            };
        }

        /// <summary>
        /// Change <see cref="PrintDialogSettings"/> property value.
        /// </summary>
        /// <param name="propertyName">The property name.</param>
        /// <param name="propertyValue">The value will change to.</param>
        /// <returns>The <see cref="PrintDialogSettings"/> after change.</returns>
        public PrintDialogSettings ChangePropertyValue(string propertyName, object propertyValue)
        {
            PropertyInfo property = this.GetType().GetProperty(propertyName, BindingFlags.Instance);

            if (property == null)
            {
                if (this.GetType().GetProperty(propertyName, BindingFlags.NonPublic) != null)
                {
                    throw new Exception("Can't change private, protected or internal property.");
                }

                throw new Exception("Can't find property. Please make sure the property name is correct.");
            }
            else
            {
                try
                {
                    property.SetValue(this, propertyValue);
                }
                catch
                {
                    throw new Exception("Can't change property's value.");
                }
            }

            return this;
        }
    }

    public class DocumentInfo
    {
        /// <summary>
        /// Initialize a <see cref="DocumentInfo"/> class.
        /// </summary>
        public DocumentInfo()
        {
            Scale = null;
            Size = null;
            Margin = null;
            Pages = null;
            PagesPerSheet = null;
            LabelQuantityFactor = null;
            Color = null;
            PageOrder = null;
            Orientation = null;
        }

        /// <summary>
        /// Page size which not calculate with orientation.
        /// </summary>
        public Size? Size { get; internal set; } = null;

        /// <summary>
        /// Pages that need to display. The array include each page count.
        /// </summary>
        public int[] Pages { get; internal set; } = null;

        /// <summary>
        /// Page scale, a percentage value, except <see cref="Double.NaN"/> means auto fit.
        /// </summary>
        public double? Scale { get; internal set; } = null;

        /// <summary>
        /// Page margin.
        /// </summary>
        public double? Margin { get; internal set; } = null;

        /// <summary>
        /// Pages per sheet.
        /// </summary>
        public int? PagesPerSheet { get; internal set; } = null;

        public double? LabelQuantityFactor { get; internal set; } = null;

        /// <summary>
        /// Color.
        /// </summary>
        public PrintSettings.PageColor? Color { get; internal set; } = null;

        /// <summary>
        /// Page order.
        /// </summary>
        public PrintSettings.PageOrder? PageOrder { get; internal set; } = null;

        /// <summary>
        /// Page orientation.
        /// </summary>
        public PrintSettings.PageOrientation? Orientation { get; internal set; } = null;
    }

    public class PrintDialog
    {
        /// <summary>
        /// Initialize a <see cref="PrintDialog"/> class.
        /// </summary>
        public PrintDialog()
        {
            Owner = null;
            Title = "标签打印";
            Icon = null; // Use the default icon
            Topmost = false;
            ShowInTaskbar = false;
            AllowScaleOption = true;
            AllowPagesOption = true;
            AllowTwoSidedOption = true;
            AllowPageOrderOption = true;
            AllowPagesPerSheetOption = true;
            AllowAddNewPrinterButton = true;
            AllowMoreSettingsExpander = true;
            AllowPrinterPreferencesButton = true;
            Document = null;
            DocumentMargin = 0;
            DocumentName = "Label Document";
            ResizeMode = ResizeMode.CanResize;
            DefaultSettings = new PrintDialogSettings();
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            GetReloadPageMethod = null;
        }

        /// <summary>
        /// The print dialog window.
        /// </summary>
        internal Internal.PrintWindow PrintWindow = null;

        /// <summary>
        /// <see cref="PrintDialog"/>'s owner.
        /// </summary>
        public Window Owner { get; set; } = null;

        /// <summary>
        /// <see cref="PrintDialog"/>'s title.
        /// </summary>
        public string Title { get; set; } = "Print";

        /// <summary>
        /// <see cref="PrintDialog"/>'s icon. Null means use default icon.
        /// </summary>
        public ImageSource Icon { get; set; } = null;

        /// <summary>
        /// The document margin info.
        /// </summary>
        public double DocumentMargin { get; set; } = 0;

        /// <summary>
        /// Allow <see cref="PrintDialog"/> at top most or not.
        /// </summary>
        public bool Topmost { get; set; } = false;

        /// <summary>
        /// Alllow <see cref="PrintDialog"/> show in taskbar or not.
        /// </summary>
        public bool ShowInTaskbar { get; set; } = false;

        /// <summary>
        /// Allow scale option or not.
        /// </summary>
        public bool AllowScaleOption { get; set; } = true;

        /// <summary>
        /// Allow pages option (contains "All Pages", "Current Page", and "Custom Pages") or not.
        /// </summary>
        public bool AllowPagesOption { get; set; } = true;

        /// <summary>
        /// Allow two-sided option or not.
        /// </summary>
        public bool AllowTwoSidedOption { get; set; } = true;

        /// <summary>
        /// Allow page order option or not if allow pages per sheet option.
        /// </summary>
        public bool AllowPageOrderOption { get; set; } = true;

        /// <summary>
        /// Allow pages per sheet option or not.
        /// </summary>
        public bool AllowPagesPerSheetOption { get; set; } = true;

        /// <summary>
        /// Allow add new printer button in printer list or not.
        /// </summary>
        public bool AllowAddNewPrinterButton { get; set; } = true;

        /// <summary>
        /// Allow more settings expander or just show all settings.
        /// </summary>
        public bool AllowMoreSettingsExpander { get; set; } = true;

        /// <summary>
        /// Allow printer preferences button or not.
        /// </summary>
        public bool AllowPrinterPreferencesButton { get; set; } = true;

        /// <summary>
        /// The document that need to print.
        /// </summary>
        public FixedDocument Document { get; set; } = null;

        /// <summary>
        /// The document name that will display in print list.
        /// </summary>
        public string DocumentName { get; set; } = "Untitled Document";

        /// <summary>
        /// <see cref="PrintDialog"/>'s resize mode.
        /// </summary>
        public ResizeMode ResizeMode { get; set; } = ResizeMode.NoResize;

        /// <summary>
        /// The default settings.
        /// </summary>
        public PrintDialogSettings DefaultSettings { get; set; } = new PrintDialogSettings();

        /// <summary>
        /// <see cref="PrintDialog"/>'s startup location.
        /// </summary>
        public WindowStartupLocation WindowStartupLocation { get; set; } = WindowStartupLocation.CenterScreen;

        /// <summary>
        /// The method that will use to get document when reload document. You can only change the content in the document. The method must return a list of <see cref="PageContent"/> that represents the page content in order.
        /// </summary>
        public Func<DocumentInfo, List<PageContent>> GetReloadPageMethod { get; set; } = null;

        /// <summary>
        /// The total sheets number that the printer will use.
        /// </summary>
        public int? TotalSheets
        {
            get
            {
                return PrintWindow.TotalSheets;
            }
        }

        /// <summary>
        /// Show <see cref="PrintDialog"/>.
        /// </summary>
        /// <exception cref="ArgumentNullException"><see cref="Document"/> or <see cref="DocumentName"/> is null.</exception>
        /// <exception cref="ArgumentException">The <see cref="DefaultSettings"/>'s pages per sheet value is invalid.</exception>
        /// <exception cref="PrintDialogExceptions.DocumentEmptyException"><see cref="Document"/> doesn't have pages.</exception>
        /// <exception cref="PrintDialogExceptions.UndefinedException"><see cref="PrintDialog"/> meet some undefined exceptions.</exception>
        /// <returns>A boolean value, true means Print button clicked, false means Cancel button clicked, null means can't open <see cref="PrintDialog"/> or there is already a running <see cref="PrintDialog"/>.</returns>
        [Obsolete]
        public bool? ShowDialog(bool v)
        {
            return ShowDialog(false, null);
        }

        /// <summary>
        /// Show <see cref="PrintDialog"/>.
        /// </summary>
        /// <param name="loading">Display loading page or not.</param>
        /// <param name="loadingAction">The method used to loading document if display loading page.</param>
        /// <exception cref="ArgumentNullException"><see cref="Document"/> or <see cref="DocumentName"/> is null.</exception>
        /// <exception cref="ArgumentException">The <see cref="DefaultSettings"/>'s pages per sheet value is invalid.</exception>
        /// <exception cref="PrintDialogExceptions.DocumentEmptyException"><see cref="Document"/> doesn't have pages.</exception>
        /// <exception cref="PrintDialogExceptions.UndefinedException"><see cref="PrintDialog"/> meet some undefined exceptions.</exception>
        /// <returns>A boolean value, true means Print button clicked, false means Cancel button clicked, null means can't open <see cref="PrintDialog"/> or there is already a running <see cref="PrintDialog"/>.</returns>
        public bool? ShowDialog(bool loading, Action loadingAction)
        {
            if (PrintWindow == null)
            {
                if (loading == false)
                {
                    if (Document == null)
                    {
                        throw new ArgumentNullException("The document is null.");
                    }
                    if (DocumentName == null)
                    {
                        throw new ArgumentNullException("The document name is null.");
                    }
                    if (Internal.PreviewHelper.MultiPagesPerSheetHelper.GetPagePerSheetCountIndex(DefaultSettings.PagesPerSheet) == -1)
                    {
                        throw new ArgumentException("The default settings's pages per sheet value is invalid. You can only set 1, 2, 4, 6, 9 or 16.");
                    }
                    if (Document.Pages.Count == 0)
                    {
                        throw new PrintDialogExceptions.DocumentEmptyException("PrintDialog can't print an empty document.", Document);
                    }
                }

                try
                {
                    PrintWindow = new Internal.PrintWindow(Document, DocumentName, DocumentMargin, DefaultSettings, AllowPagesOption, AllowScaleOption, AllowTwoSidedOption, AllowPagesPerSheetOption, AllowPageOrderOption, AllowMoreSettingsExpander, AllowAddNewPrinterButton, AllowPrinterPreferencesButton, GetReloadPageMethod, loading, loadingAction)
                    {
                        Title = Title,
                        Owner = Owner,
                        Topmost = Topmost,
                        ResizeMode = ResizeMode,
                        ShowInTaskbar = ShowInTaskbar,
                        WindowStartupLocation = WindowStartupLocation
                    };

                    if (Icon != null)
                    {
                        PrintWindow.Icon = Icon;
                    }

                    PrintWindow.ShowDialog();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    //throw new PrintDialogExceptions.UndefinedException(ex.Message, ex);
                }

                bool returnValue = PrintWindow.ReturnValue;

                PrintWindow = null;

                return returnValue;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Switch the current running <see cref="PrintDialog"/>'s page into settings and preview page.
        /// </summary>
        /// <exception cref="ArgumentNullException"><see cref="Document"/> or <see cref="DocumentName"/> is null.</exception>
        /// <exception cref="ArgumentException">The <see cref="DefaultSettings"/>'s pages per sheet value is invalid.</exception>
        /// <exception cref="PrintDialogExceptions.DocumentEmptyException"><see cref="Document"/> doesn't have pages.</exception>
        /// <exception cref="PrintDialogExceptions.UndefinedException"><see cref="PrintDialog"/> meet some undefined exceptions.</exception>
        public void LoadingEnd()
        {
            if (PrintWindow != null)
            {
                if (Document == null)
                {
                    throw new ArgumentNullException("The document is null.");
                }
                if (DocumentName == null)
                {
                    throw new ArgumentNullException("The document name is null.");
                }
                if (Internal.PreviewHelper.MultiPagesPerSheetHelper.GetPagePerSheetCountIndex(DefaultSettings.PagesPerSheet) == -1)
                {
                    throw new ArgumentException("The default settings's pages per sheet value is invalid. You can only set 1, 2, 4, 6, 9 or 16.");
                }
                if (Document.Pages.Count == 0)
                {
                    throw new PrintDialogExceptions.DocumentEmptyException("PrintDialog can't print an empty document.", Document);
                }

                PrintWindow.BeginSettingAndPreviewing(Document, DocumentName, DocumentMargin, DefaultSettings, AllowPagesOption, AllowScaleOption, AllowTwoSidedOption, AllowPagesPerSheetOption, AllowPageOrderOption, AllowMoreSettingsExpander, AllowAddNewPrinterButton, AllowPrinterPreferencesButton, GetReloadPageMethod);
            }
        }
    }
}
