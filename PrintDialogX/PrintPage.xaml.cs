using System;
using System.IO;
using System.IO.Packaging;
using System.Runtime.InteropServices;
using System.Printing;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Xps;
using System.Windows.Xps.Packaging;
using System.Windows.Media;
using System.Windows.Markup;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Threading;
using System.Linq;
using System.Xml.Linq;
using System.Text;
using System.Security.Cryptography;
using System.Xml;
using static PrintDialogX.Setting;
using Image = System.Windows.Controls.Image;
using Orientation = System.Windows.Controls.Orientation;
using MessageBox = System.Windows.MessageBox;
using Size = System.Windows.Size;
using Point = System.Windows.Point;
using ComboBox = System.Windows.Controls.ComboBox;
using PrintDialogX.Common;

namespace PrintDialogX.Internal
{
    /// <summary>
    /// The PrintDialog page.
    /// </summary>
    partial class PrintPage : Page
    {
        //Properties

        #region Public properties

        /// <summary>
        /// A boolean value, true means Print button clicked, false means Cancel button clicked.
        /// </summary>
        public bool ReturnValue { get; internal set; } = false;

        /// <summary>
        /// The total sheets number that the printer will use.
        /// </summary>
        public int? TotalSheets
        {
            get
            {
                try
                {
                    if (twoSidedCheckBox.IsChecked == false)
                    {
                        return documentPreviewer.PageCount * int.Parse(copiesNumberPicker.Text);
                    }
                    else
                    {
                        return (int)Math.Ceiling(documentPreviewer.PageCount * int.Parse(copiesNumberPicker.Text) / 2.0);
                    }
                }
                catch
                {
                    return null;
                }
            }
        }

        #endregion

        #region Private properties

        private bool isLoaded;
        private bool isRefresh;
        private Package package;
        private int lastPrinterIndex;
        private List<int> lastPageList;
        private FixedDocument fixedDocument;
        //save printer setting in ini file
        private bool firstInit = true;
        private int printerID;

        private readonly Uri _xpsUrl;
        private readonly int _pageCount;
        private readonly int[] _zoomList;
        private readonly double _pageMargin;
        private readonly string _documentName;
        private readonly System.Windows.Size _orginalPageSize;
        private readonly bool _allowPagesOption;
        private readonly bool _allowScaleOption;
        private readonly bool _allowTwoSidedOption;
        private readonly bool _allowPageOrderOption;
        private readonly bool _allowPagesPerSheetOption;
        private readonly bool _allowMoreSettingsExpander;
        private readonly bool _allowAddNewPrinerComboBoxItem;
        private readonly bool _allowPrinterPreferencesButton;
        private readonly PrintServer _localDefaultPrintServer;
        private readonly List<PageContent> _orginalPagesContentList;
        private readonly PrintDialog.PrintDialogSettings 
            _defaultSettings;
        private readonly System.Windows.Controls.PrintDialog 
            _systemPrintDialog;
        private readonly Func<PrintDialog.DocumentInfo, List<PageContent>> 
            _getDocumentWhenReloadDocumentMethod;

        private static readonly DispatcherOperationCallback 
            _exitFrameCallback = new DispatcherOperationCallback(ExitFrame);
        private List<PageSize> SizeList;
        private PrintQueue _printer;

        #endregion

        //Dll Import

        #region Printer Preferences Dialog

        [DllImport("winspool.Drv", EntryPoint = "DocumentPropertiesW", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        static extern int DocumentProperties(IntPtr hwnd, IntPtr hPrinter,
        [MarshalAs(UnmanagedType.LPWStr)] string pDeviceName,
        IntPtr pDevModeOutput, ref IntPtr pDevModeInput, int fMode);
        [DllImport("kernel32.dll")]
        static extern IntPtr GlobalLock(IntPtr hMem);
        [DllImport("kernel32.dll")]
        static extern bool GlobalUnlock(IntPtr hMem);
        [DllImport("kernel32.dll")]
        static extern bool GlobalFree(IntPtr hMem);

        private void OpenPrinterPropertiesDialog(System.Drawing.Printing.PrinterSettings printerSettings)
        {
            Window dialog = new Window();

            IntPtr handle = new System.Windows.Interop.WindowInteropHelper(dialog).Handle;
            IntPtr hDevMode = printerSettings.GetHdevmode(printerSettings.DefaultPageSettings);
            IntPtr pDevMode = GlobalLock(hDevMode);
            int sizeNeeded = DocumentProperties(handle, IntPtr.Zero, printerSettings.PrinterName, pDevMode, ref pDevMode, 0);
            IntPtr devModeData = Marshal.AllocHGlobal(sizeNeeded);
            DocumentProperties(handle, IntPtr.Zero, printerSettings.PrinterName, devModeData, ref pDevMode, 14);
            GlobalUnlock(hDevMode);
            printerSettings.SetHdevmode(devModeData);
            printerSettings.DefaultPageSettings.SetHdevmode(devModeData);
            GlobalFree(hDevMode);
            Marshal.FreeHGlobal(devModeData);
        }

        #endregion

        //Initialize

        #region Initialize the window

        /// <summary>
        /// Initialize PrintDialog.
        /// </summary>
        /// <param name="document">The document that need to print.</param>
        /// <param name="documentName">The document name that will display in print list.</param>
        /// <param name="pageMargin">The page margin info.</param>
        /// <param name="defaultSettings">The default settings.</param>
        /// <param name="allowPagesOption">Allow pages option or not.</param>
        /// <param name="allowScaleOption">Allow scale option or not.</param>
        /// <param name="allowTwoSidedOption">Allow two-sided option or not.</param>
        /// <param name="allowPagesPerSheetOption">Allow pages per sheet option or not.</param>
        /// <param name="allowPageOrderOption">Allow page order option or not if allow pages per sheet option.</param>
        /// <param name="allowMoreSettingsExpander">Allow more settings expander or just show all settings.</param>
        /// <param name="allowAddNewPrinerComboBoxItem">Allow add new printer button in printer list or not.</param>
        /// <param name="allowPrinterPreferencesButton">Allow printer preferences button or not.</param>
        /// <param name="getDocumentWhenReloadDocumentMethod">The method that will use to get document when reload document. You can only change the content in the document. The method must return a list of PageContent that include the page content in order.</param>
        public PrintPage(FixedDocument document, string documentName, double pageMargin, PrintDialog.PrintDialogSettings defaultSettings, bool allowPagesOption, bool allowScaleOption, bool allowTwoSidedOption, bool allowPagesPerSheetOption, bool allowPageOrderOption, bool allowMoreSettingsExpander, bool allowAddNewPrinerComboBoxItem, bool allowPrinterPreferencesButton, Func<PrintDialog.DocumentInfo, List<PageContent>> getDocumentWhenReloadDocumentMethod)
        {
            isLoaded = false;

            InitializeComponent();

            this.UpdateLayout();
            DoEvents();

            isRefresh = true;
            lastPageList = new List<int>();
            fixedDocument = document;

            _pageMargin = pageMargin;
            _documentName = documentName;
            _defaultSettings = defaultSettings;
            _allowPagesOption = allowPagesOption;
            _allowScaleOption = allowScaleOption;
            _allowTwoSidedOption = allowTwoSidedOption;
            _allowPageOrderOption = allowPageOrderOption;
            _allowPagesPerSheetOption = allowPagesPerSheetOption;
            _allowMoreSettingsExpander = allowMoreSettingsExpander;
            _allowAddNewPrinerComboBoxItem = allowAddNewPrinerComboBoxItem;
            _allowPrinterPreferencesButton = allowPrinterPreferencesButton;
            _getDocumentWhenReloadDocumentMethod = getDocumentWhenReloadDocumentMethod;

            _pageCount = document.Pages.Count;
            _orginalPageSize = document.DocumentPaginator.PageSize;

            _xpsUrl = new Uri("memorystream://" + Guid.NewGuid().ToString() + ".xps");
            _zoomList = new int[] { 25, 50, 75, 100, 150, 200 };
            _localDefaultPrintServer = new PrintServer();
            _orginalPagesContentList = new List<PageContent>();
            _systemPrintDialog = new System.Windows.Controls.PrintDialog();

            foreach (PageContent page in document.Pages)
            {
                string xaml = XamlWriter.Save(page);
                PageContent pageClone = XamlReader.Parse(xaml) as PageContent;

                _orginalPagesContentList.Add(pageClone);
            }
        }

        #endregion

        //Methods

        #region Define DoEvents() method to refresh the interface

        internal static void DoEvents()
        {
            DispatcherFrame nestedFrame = new DispatcherFrame();
            DispatcherOperation exitOperation = Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background, _exitFrameCallback, nestedFrame);
            Dispatcher.PushFrame(nestedFrame);

            if (exitOperation.Status != DispatcherOperationStatus.Completed)
            {
                exitOperation.Abort();
            }
        }

        private static object ExitFrame(object state)
        {
            DispatcherFrame frame = state as DispatcherFrame;
            frame.Continue = false;
            return null;
        }

        #endregion

        #region Define GetChildren() method to get all chilren of control

        internal static List<DependencyObject> GetChildren(DependencyObject element)
        {
            List<DependencyObject> elements = new List<DependencyObject>();

            if (VisualTreeHelper.GetChildrenCount(element) > 0)
            {
                for (int i = 0; i < VisualTreeHelper.GetChildrenCount(element); i++)
                {
                    DependencyObject control = VisualTreeHelper.GetChild(element, i);
                    if (control != null)
                    {
                        elements.Add(control);
                        elements.AddRange(GetChildren(control));
                    }
                }
            }

            return elements;
        }

        #endregion

        #region Define GetCurrentPage() method to get the page index which user is looking

        private int GetCurrentPage()
        {
            int page;
            ScrollViewer scrollViewer = (ScrollViewer)documentPreviewer.Template.FindName("PART_ContentHost", documentPreviewer);

            if (documentPreviewer.MaxPagesAcross == 1)
            {
                page = (int)(scrollViewer.VerticalOffset / (scrollViewer.ExtentHeight / documentPreviewer.PageCount)) + 1;
            }
            else
            {
                page = ((int)(scrollViewer.VerticalOffset / (scrollViewer.ExtentHeight / Math.Ceiling((double)documentPreviewer.PageCount / 2))) + 1) * 2 - 1;
            }

            if (page < 1)
            {
                page = 1;
            }
            if (page > documentPreviewer.PageCount)
            {
                page = documentPreviewer.PageCount;
            }

            return page;
        }

        #endregion

        #region Define LoadPrinters() method to find all printers

        private void LoadPrinters()
        {
            try
            {
                int equipmentComboBoxSelectedIndex = 0;
                printerComboBox.Items.Clear();

                if (_allowAddNewPrinerComboBoxItem == true)
				{
					try
					{
						Grid itemMainGrid = new Grid();
						itemMainGrid.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(55) });
						itemMainGrid.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(10) });
						itemMainGrid.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(1, GridUnitType.Star) });

						TextBlock textInfoBlock = new TextBlock()
						{
							Text = "添加打印机",
							FontSize = 14,
							VerticalAlignment = VerticalAlignment.Center
						};

						Image itemIcon = new Image()
						{
							Width = 55,
							Height = 55,
							Source = new System.Windows.Media.Imaging.BitmapImage(new Uri("/PrintDialog;component/Resources/AddPrinter.png", UriKind.Relative)),
							Stretch = Stretch.Fill
						};

						itemIcon.SetValue(Grid.ColumnProperty, 0);
						textInfoBlock.SetValue(Grid.ColumnProperty, 2);

						itemMainGrid.Children.Add(itemIcon);
						itemMainGrid.Children.Add(textInfoBlock);

						ComboBoxItem item = new ComboBoxItem
						{
							Content = itemMainGrid,
							Height = 55,
							ToolTip = "添加打印机"
						};

						printerComboBox.Items.Insert(0, item);
					}
					catch (Exception ex)
					{
                        //MessageBox.Show(ex.Message);
					}
                }

                foreach (PrintQueue printer in _localDefaultPrintServer.GetPrintQueues())
                {
					try
					{
						printer.Refresh();

						string status = PrinterHelper.PrinterHelper.GetPrinterStatusInfo(_localDefaultPrintServer.GetPrintQueue(printer.FullName));
						string loction = printer.Location;
						string comment = printer.Comment;

						if (String.IsNullOrWhiteSpace(loction) == true)
						{
							loction = "Unknown";
						}
						if (String.IsNullOrWhiteSpace(comment) == true)
						{
							comment = "Unknown";
						}

						Grid itemMainGrid = new Grid();
						itemMainGrid.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(55) });
						itemMainGrid.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(10) });
						itemMainGrid.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(1, GridUnitType.Star) });

						StackPanel textInfoPanel = new StackPanel()
						{
							Orientation = Orientation.Vertical,
							VerticalAlignment = VerticalAlignment.Center
						};
						textInfoPanel.Children.Add(new TextBlock() { Text = printer.FullName, FontSize = 14 });
						textInfoPanel.Children.Add(new TextBlock() { Text = status, Margin = new Thickness(0, 7, 0, 0) });

						Image printerIcon = new Image()
						{
							Width = 55,
							Height = 55,
							Source = PrinterInfoHelper.PrinterIconHelper.GetPrinterIcon(printer, _localDefaultPrintServer),
							Stretch = Stretch.Fill
						};

						if (printer.IsOffline == true)
						{
							printerIcon.Opacity = 0.5;
						}

						printerIcon.SetValue(Grid.ColumnProperty, 0);
						textInfoPanel.SetValue(Grid.ColumnProperty, 2);

						itemMainGrid.Children.Add(printerIcon);
						itemMainGrid.Children.Add(textInfoPanel);

						ComboBoxItem item = new ComboBoxItem
						{
							Content = itemMainGrid,
							Height = 55,
							ToolTip = "Name: " + printer.FullName + "\nStatus: " + status + "\nDocument: " + printer.NumberOfJobs + "\nLoction: " + loction + "\nComment: " + comment,
							Tag = printer.FullName
						};

						printerComboBox.Items.Insert(0, item);

						if (LocalPrintServer.GetDefaultPrintQueue().FullName == printer.FullName)
						{
							equipmentComboBoxSelectedIndex = printerComboBox.Items.Count;
						}
					}
					catch (Exception ex)
					{
                        // skip printer
                        //MessageBox.Show(ex.Message);
					}
                }

                if (printerComboBox.Items.Count == 0)
                {
                    MessageWindow window = new MessageWindow("There is no printer support.", "Error", "OK", null, MessageWindow.MessageIcon.Error)
                    {
                        Owner = Window.GetWindow(this),
                        Icon = Window.GetWindow(this).Icon
                    };
                    window.ShowDialog();

                    Window.GetWindow(this).Close();
                }

				//// test
				//equipmentComboBoxSelectedIndex = equipmentComboBox.Items.Count - equipmentComboBoxSelectedIndex;
				//lastPrinterIndex = equipmentComboBoxSelectedIndex;
				//equipmentComboBox.SelectedIndex = equipmentComboBoxSelectedIndex;
				//equipmentComboBox.Tag = PrinterInfoHelper.PrinterIconHelper.GetPrinterIcon(_localDefaultPrintServer.GetPrintQueue((equipmentComboBox.SelectedItem as ComboBoxItem).Tag.ToString()), _localDefaultPrintServer, true);
				//equipmentComboBox.Text = (equipmentComboBox.SelectedItem as ComboBoxItem).Tag.ToString();

                // test
				try
				{
					equipmentComboBoxSelectedIndex = printerComboBox.Items.Count - equipmentComboBoxSelectedIndex;
				}
				catch (Exception ex)
				{
					ShowMessage("1: " + ex.Message);
				}

				try
				{
					lastPrinterIndex = equipmentComboBoxSelectedIndex;
				}
				catch (Exception ex)
				{
					ShowMessage("2: " + ex.Message);
				}
				try
				{
					printerComboBox.SelectedIndex = equipmentComboBoxSelectedIndex;
				}
				catch (Exception ex)
				{
					ShowMessage("3: " + ex.Message);
				}
				try
				{//todo watch
					printerComboBox.Tag = PrinterInfoHelper.PrinterIconHelper.GetPrinterIcon(_localDefaultPrintServer.GetPrintQueue((printerComboBox.SelectedItem as ComboBoxItem).Tag.ToString()), _localDefaultPrintServer, true);
				}
				catch (Exception ex)
				{
					ShowMessage("4: " + ex.Message);
				}
				try
				{
					printerComboBox.Text = (printerComboBox.SelectedItem as ComboBoxItem).Tag.ToString();
				}
				catch (Exception ex)
				{
					ShowMessage("5: " + ex.Message);
				}
			}
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        #endregion

        #region Define LoadPrinterSettings() method to load the printer's settings

        private void LoadPrinterSettings(bool useDefaults = true)
        {
            string lastColor = "";
            string lastQuality = "";
            string lastSize = "";
            string lastType = "";
            string lastSource = "";

            if (useDefaults == false)
            {
                try
                {
                    lastColor = (colorComboBox.SelectedItem as ComboBoxItem).Content.ToString();
                    lastQuality = (qualityComboBox.SelectedItem as ComboBoxItem).Content.ToString();
                    //lastSize = (sizeComboBox.SelectedItem as ComboBoxItem).Content.ToString();
                    lastType = (typeComboBox.SelectedItem as ComboBoxItem).Content.ToString();
                    lastSource = (sourceComboBox.SelectedItem as ComboBoxItem).Content.ToString();
                }
                catch
                {
                    lastColor = _defaultSettings.Color.ToString();
                    lastQuality = _defaultSettings.Quality.ToString();
                    //lastSize = PrintSettings.SettingsHepler.NameInfoHepler.GetPageMediaSizeNameInfo(_defaultSettings.PageSize);
                    lastType = PrintSettings.SettingsHepler.NameInfoHepler.GetPageMediaTypeNameInfo(_defaultSettings.PageType);
                    lastSource = PrintSettings.SettingsHepler.NameInfoHepler.GetInputBinNameInfo(InputBin.AutoSelect);
                }
            }

            _printer = _localDefaultPrintServer.GetPrintQueue((printerComboBox.SelectedItem as ComboBoxItem).Tag.ToString());

            if (_printer.GetPrintCapabilities().MaxCopyCount.HasValue)
            {
                copiesNumberPicker.MaxValue = _printer.GetPrintCapabilities().MaxCopyCount.Value;
            }
            else
            {
                copiesNumberPicker.MaxValue = 1;
            }

            // copies
            copiesNumberPicker.ChangeButtonEnabled();
            if (int.Parse(copiesNumberPicker.Text) > 1)
            {
                collateCheckBox.Visibility = Visibility.Visible;
            }
            else
            {
                collateCheckBox.Visibility = Visibility.Collapsed;
            }

            // Color
            int colorComboBoxSelectedIndex = 0;
            colorComboBox.Items.Clear();
            foreach (OutputColor color in _printer.GetPrintCapabilities().OutputColorCapability)
            {
                ComboBoxItem item = new ComboBoxItem
                {
                    Content = color.ToString()
                };

                colorComboBox.Items.Add(item);

                if (useDefaults == true)
                {
                    if (_defaultSettings.UsePrinterDefaultSettings == false)
                    {
                        if (color.ToString() == _defaultSettings.Color.ToString())
                        {
                            colorComboBoxSelectedIndex = colorComboBox.Items.Count - 1;
                        }
                    }
                    else
                    {
                        if (color == _printer.DefaultPrintTicket.OutputColor)
                        {
                            colorComboBoxSelectedIndex = colorComboBox.Items.Count - 1;
                        }
                    }
                }
                else
                {
                    if (color.ToString() == lastColor)
                    {
                        colorComboBoxSelectedIndex = colorComboBox.Items.Count - 1;
                    }
                }
            }
            if (colorComboBox.HasItems == false)
            {
                ComboBoxItem item = new ComboBoxItem
                {
                    Content = _defaultSettings.Color.ToString()
                };

                colorComboBox.Items.Add(item);
            }
            colorComboBox.SelectedIndex = colorComboBoxSelectedIndex;

            // Quality
            int qualityComboBoxSelectedIndex = 0;
            qualityComboBox.Items.Clear();
            foreach (OutputQuality quality in _printer.GetPrintCapabilities().OutputQualityCapability)
            {
                ComboBoxItem item = new ComboBoxItem
                {
                    Content = quality.ToString()
                };

                qualityComboBox.Items.Add(item);

                if (useDefaults == true)
                {
                    if (_defaultSettings.UsePrinterDefaultSettings == false)
                    {
                        if (quality.ToString() == _defaultSettings.Quality.ToString())
                        {
                            qualityComboBoxSelectedIndex = qualityComboBox.Items.Count - 1;
                        }
                    }
                    else
                    {
                        if (quality == _printer.DefaultPrintTicket.OutputQuality)
                        {
                            qualityComboBoxSelectedIndex = qualityComboBox.Items.Count - 1;
                        }
                    }
                }
                else
                {
                    if (quality.ToString() == lastQuality)
                    {
                        qualityComboBoxSelectedIndex = qualityComboBox.Items.Count - 1;
                    }
                }
            }
            if (qualityComboBox.HasItems == false)
            {
                ComboBoxItem item = new ComboBoxItem
                {
                    Content = _defaultSettings.Quality.ToString()
                };

                qualityComboBox.Items.Add(item);
            }
            qualityComboBox.SelectedIndex = qualityComboBoxSelectedIndex;

            // Twosided
            if (_defaultSettings.UsePrinterDefaultSettings == true)
            {
                if (_printer.DefaultPrintTicket.Duplexing.HasValue)
                {
                    if (_printer.DefaultPrintTicket.Duplexing.Value == Duplexing.TwoSidedLongEdge)
                    {
                        twoSidedCheckBox.IsChecked = true;
                        twoSidedTypeComboBox.SelectedIndex = 0;
                    }
                    else if (_printer.DefaultPrintTicket.Duplexing.Value == Duplexing.TwoSidedShortEdge)
                    {
                        twoSidedCheckBox.IsChecked = true;
                        twoSidedTypeComboBox.SelectedIndex = 1;
                    }
                    else
                    {
                        twoSidedCheckBox.IsChecked = false;
                    }
                }
                else
                {
                    twoSidedCheckBox.IsChecked = false;
                }
            }

            // Page size
            SizeList = GetSizeList(_printer);

            FillSizeComboBox(SizeList, useDefaults);

            // Page media type: plain
            int typeComboBoxSelectedIndex = 0;
            typeComboBox.Items.Clear();
            foreach (PageMediaType type in _printer.GetPrintCapabilities().PageMediaTypeCapability)
            {
                ComboBoxItem item = new ComboBoxItem
                {
                    Content = PrintSettings.SettingsHepler.NameInfoHepler.GetPageMediaTypeNameInfo(type)
                };

                typeComboBox.Items.Add(item);

                if (useDefaults == true)
                {
                    if (_defaultSettings.UsePrinterDefaultSettings == false)
                    {
                        if (type.ToString() == _defaultSettings.PageType.ToString())
                        {
                            typeComboBoxSelectedIndex = typeComboBox.Items.Count - 1;
                        }
                    }
                    else
                    {
                        if (type == _printer.DefaultPrintTicket.PageMediaType)
                        {
                            typeComboBoxSelectedIndex = typeComboBox.Items.Count - 1;
                        }
                    }
                }
                else
                {
                    if (PrintSettings.SettingsHepler.NameInfoHepler.GetPageMediaTypeNameInfo(type) == lastType)
                    {
                        typeComboBoxSelectedIndex = typeComboBox.Items.Count - 1;
                    }
                }
            }
            if (typeComboBox.HasItems == false)
            {
                ComboBoxItem item = new ComboBoxItem
                {
                    Content = PrintSettings.SettingsHepler.NameInfoHepler.GetPageMediaTypeNameInfo(_defaultSettings.PageType)
                };

                typeComboBox.Items.Add(item);
            }
            typeComboBox.SelectedIndex = typeComboBoxSelectedIndex;

            // Input bin
            int sourceComboBoxSelectedIndex = 0;
            sourceComboBox.Items.Clear();
            foreach (InputBin source in _printer.GetPrintCapabilities().InputBinCapability)
            {
                ComboBoxItem item = new ComboBoxItem
                {
                    Content = PrintSettings.SettingsHepler.NameInfoHepler.GetInputBinNameInfo(source)
                };

                sourceComboBox.Items.Add(item);

                if (useDefaults == true)
                {
                    if (source == _printer.DefaultPrintTicket.InputBin)
                    {
                        sourceComboBoxSelectedIndex = sourceComboBox.Items.Count - 1;
                    }
                }
                else
                {
                    if (PrintSettings.SettingsHepler.NameInfoHepler.GetInputBinNameInfo(source) == lastSource)
                    {
                        sourceComboBoxSelectedIndex = sourceComboBox.Items.Count - 1;
                    }
                }
            }
            if (sourceComboBox.HasItems == false)
            {
                ComboBoxItem item = new ComboBoxItem
                {
                    Content = PrintSettings.SettingsHepler.NameInfoHepler.GetInputBinNameInfo(InputBin.AutoSelect)
                };

                sourceComboBox.Items.Add(item);
            }
            sourceComboBox.SelectedIndex = sourceComboBoxSelectedIndex;
        }

		private void FillSizeComboBox(List<PageSize> SizeList, bool useDefaults)
		{
			int sizeComboBoxSelectedIndex = 0;
            sizeComboBox.Items.Clear();
			foreach (PageSize size in SizeList)
            {
                ComboBoxItem item = new ComboBoxItem
                {
                    Content = size.Name 
                    + ": " 
                    + size.widthStr.Substring(0, size.widthStr.Length - 3)
                    + " x "
                    + size.heightStr.Substring(0, size.heightStr.Length - 3)
                    + " mm"
                };

                sizeComboBox.Items.Add(item);

                if (useDefaults == true)
                {
                    if (_defaultSettings.UsePrinterDefaultSettings == false)
                    {
                        if (size.NameID == _defaultSettings.NameID)
                        {
                            sizeComboBoxSelectedIndex = sizeComboBox.Items.Count - 1;
                        }
                    }
                    else
                    {
                        if (size.NameID == 0)//size == printer.DefaultPrintTicket.PageMediaSize
                        {
                            sizeComboBoxSelectedIndex = sizeComboBox.Items.Count - 1;
                        }
                    }
                }
                else
                {
					//todo check
					//if (PrintSettings.SettingsHepler.NameInfoHepler.GetPageMediaSizeNameInfo(size.PageMediaSizeName.Value) == lastSize)
					//{
					//	sizeComboBoxSelectedIndex = sizeComboBox.Items.Count - 1;
					//}
				}
            }

            if (sizeComboBox.HasItems == false)
            {
                //todo check
                ComboBoxItem item = new ComboBoxItem
                {
                    Content = PrintSettings.SettingsHepler.NameInfoHepler.GetPageMediaSizeNameInfo(PageMediaSizeName.ISOA4)//_defaultSettings.PageSize)
                };

                sizeComboBox.Items.Add(item);
            }

            sizeComboBox.SelectedIndex = sizeComboBoxSelectedIndex;
		}

		private void ShowMessage(string message)
		{
			MessageBox.Show(message);
		}

		private List<PageSize> GetSizeList(PrintQueue printer)
		{
			List<PageSize> SizeList = new List<PageSize>();

		    var xmlDoc = new XmlDocument();
			using (MemoryStream printerCapXmlStream = printer.GetPrintCapabilitiesAsXml())
			{
				xmlDoc.Load(printerCapXmlStream);
			}
			
            var manager = new XmlNamespaceManager(xmlDoc.NameTable);
			manager.AddNamespace(xmlDoc.DocumentElement.Prefix, xmlDoc.DocumentElement.NamespaceURI);
			var nodeList = xmlDoc.SelectNodes("//psf:Feature[@name='psk:PageMediaSize']/psf:Option", manager);

            //( 800mm) / (25.4 mm/in) / (1/96in) = 3023.62
            //double rate = 1 / 1000 / 25.4 / (1 / 96);
                //270; // 96 / 25.4 * 100; // 264.583 / 1.041
            //todo test
            //int SH = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height;
          //int SW = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width;

			for (int i = 0; i < nodeList.Count; i++)
			{
				try
				{
					XmlNode node = nodeList[i];
					PageSize size = new PageSize();
					size.NameID = i;
					size.Width = Convert.ToDouble(node.ChildNodes[0].InnerText) / 1000 * (96 / 25.4);
					size.Height = Convert.ToDouble(node.ChildNodes[1].InnerText) / 1000 * (96 / 25.4);
					size.Name = node.LastChild.InnerText;

                    // original value
                    size.nameStr = node.Attributes["name"].Value;
                    size.widthStr = node.ChildNodes[0].InnerText;
                    size.heightStr = node.ChildNodes[1].InnerText;

					SizeList.Add(size);
				}
				catch (Exception)
				{
					continue;
				}
			}
            return SizeList;
		}

		#endregion

		#region Define LoadPreview() method to make XPS document and load document preview

		private void LoadPreview(FixedDocument doc)
        {
            PackageStore.RemovePackage(_xpsUrl);
            MemoryStream stream = new MemoryStream();
            package = Package.Open(stream, FileMode.OpenOrCreate, FileAccess.ReadWrite);

            PackageStore.AddPackage(_xpsUrl, package);
            XpsDocument xpsDoc = new XpsDocument(package);
            try
            {
                xpsDoc.Uri = _xpsUrl;
                XpsDocumentWriter writer = XpsDocument.CreateXpsDocumentWriter(xpsDoc);
                writer.Write(((IDocumentPaginatorSource)doc).DocumentPaginator);

                documentPreviewer.Document = xpsDoc.GetFixedDocumentSequence();
            }
            finally
            {
                if (xpsDoc != null)
                {
                    xpsDoc.Close();
                }
            }
        }

        #endregion

        #region Define ReloadDocument() method to reload document with user settings

        private void ReloadDocument()
        {
            if (isRefresh == true)
            {
                loadingGrid.Visibility = Visibility.Visible;
                DoEvents();
                try
                {
					#region Printer and Page Size
					PageSize size;
					printerID = printerComboBox.SelectedIndex;
					_printer = _localDefaultPrintServer.GetPrintQueue((printerComboBox.SelectedItem as ComboBoxItem).Tag.ToString());
					SizeList = GetSizeList(_printer);
                    if (SizeList.Count > 0)
                    {
                        int sizeID;
                        // first init && using system default printer
                        if (firstInit)
                        {
                            // check ini file
                            if (!File.Exists(g_ConfigPath))
                            {
                                FileHelper.CreateFileReplace(g_ConfigPath, "");
                            }
                                

                            string printerIDStr = IniHelper.GetIniData("Config", "PrinterID", g_ConfigPath);
                            string sizeIDStr = IniHelper.GetIniData("Config", "PrintPageSizeID", g_ConfigPath);
                            if (printerIDStr != null && printerIDStr.Length > 0 
                                && sizeIDStr != null && sizeIDStr.Length > 0)
                            {
                                int printerBoxIDOrigin = printerComboBox.SelectedIndex;
                                printerComboBox.SelectedIndex = Convert.ToInt32(printerIDStr);
								try
								{
                                    //check new printer
                                    PrintQueue newPrinter = _localDefaultPrintServer.GetPrintQueue((printerComboBox.SelectedItem as ComboBoxItem).Tag.ToString());
								    List<PageSize> NewSizeList = GetSizeList(newPrinter);
                                    if (NewSizeList.Count > 0) // Change Printer
                                    {
                                        _printer = newPrinter;
                                        FillSizeComboBox(NewSizeList, false);
                                        sizeID = Convert.ToInt32(sizeIDStr);
                                        sizeComboBox.SelectedIndex = sizeID;
                                    }
                                    else //use original
                                    { 
                                        printerComboBox.SelectedIndex = printerBoxIDOrigin;
                                        sizeID = sizeComboBox.SelectedIndex;
                                    }
								}
								catch (Exception)//use original
								{
									printerComboBox.SelectedIndex = printerBoxIDOrigin;
									sizeID = sizeComboBox.SelectedIndex;
								}
							}
                            else// save default
                            { 
                                sizeID = sizeComboBox.SelectedIndex;
                                IniHelper.SaveIniData("Config", "PrinterID", printerID.ToString(), g_ConfigPath);
                                IniHelper.SaveIniData("Config", "PrintPageSizeID", sizeID.ToString(), g_ConfigPath);
                            }
                            firstInit = false;
                        }
                        else// by selection, save selection
                        { 
                            sizeID = sizeComboBox.SelectedIndex;
                            IniHelper.SaveIniData("Config", "PrinterID", printerID.ToString(), g_ConfigPath);
                            IniHelper.SaveIniData("Config", "PrintPageSizeID", sizeID.ToString(), g_ConfigPath);
                        }
						size = SizeList[sizeID];
                    }
                    else // No size list
                    {
                        //todo check

                        //PageMediaSizeName sizeName = PageMediaSizeName.Unknown;//.ISOA4
                        //int index = 0;
                        //foreach (string name in sizeName.GetType().GetEnumNames())
                        //{
                        //    if (name == _defaultSettings.PageSize.ToString())
                        //        sizeName = (PageMediaSizeName)index;
                        //    index++;
                        //}
                        //size = new PageMediaSize(sizeName);

                        size = new PageSize();
                        //todo fill
                    }
					#endregion
					// Load document
					FixedDocument doc = new FixedDocument();
					#region Orientation
                    if (orientationComboBox.SelectedIndex == 0)
                        doc.DocumentPaginator.PageSize = new Size(size.Width, size.Height);
                    else
                        doc.DocumentPaginator.PageSize = new Size(size.Height, size.Width);
					#endregion

					#region Page List
					List<int> pageList = new List<int>();
                    bool isCustomPages = false;
                    if (pagesComboBox.SelectedIndex == 1)
                    {
                        isCustomPages = true;

                        if (lastPageList.Count > 0)
                            pageList.Add(lastPageList[GetCurrentPage() - 1]);
                        else
                            pageList.Add(GetCurrentPage());
                    }
                    else if (pagesComboBox.SelectedIndex == 2 && String.IsNullOrWhiteSpace(customPagesTextBox.Text) == false)
                    {
                        isCustomPages = true;

                        string[] customPageInputList = customPagesTextBox.Text.Split(',');
                        foreach (string str in customPageInputList)
                        {
                            if (str.Contains("-"))
                            {
                                try
                                {
                                    string[] pageRange = str.Split('-');

                                    if (pageRange.Length != 2 || int.Parse(pageRange[0]) > int.Parse(pageRange[1]) || int.Parse(pageRange[0]) <= 0 || int.Parse(pageRange[1]) > _pageCount)
                                    {
                                        MessageWindow window = new MessageWindow("The custom pages value is invalid.", "Error", "OK", null, MessageWindow.MessageIcon.Error)
                                        {
                                            Owner = Window.GetWindow(this),
                                            Icon = Window.GetWindow(this).Icon
                                        };
                                        window.ShowDialog();

                                        pageList = new List<int>();
                                        isCustomPages = false;
                                        break;
                                    }
                                    else
                                    {
                                        for (int i = int.Parse(pageRange[0]); i <= int.Parse(pageRange[1]); i++)
                                        {
                                            pageList.Add(i);
                                        }
                                    }
                                }
                                catch
                                {
                                    MessageWindow window = new MessageWindow("The custom pages value is invalid.", "Error", "OK", null, MessageWindow.MessageIcon.Error)
                                    {
                                        Owner = Window.GetWindow(this),
                                        Icon = Window.GetWindow(this).Icon
                                    };
                                    window.ShowDialog();

                                    pageList = new List<int>();
                                    isCustomPages = false;
                                    break;
                                }
                            }
                            else
                            {
                                try
                                {
                                    if (int.Parse(str) <= 0 || int.Parse(str) > _pageCount)
                                    {
                                        MessageWindow window = new MessageWindow("The custom pages value is invalid.", "Error", "OK", null, MessageWindow.MessageIcon.Error)
                                        {
                                            Owner = Window.GetWindow(this),
                                            Icon = Window.GetWindow(this).Icon
                                        };
                                        window.ShowDialog();

                                        pageList = new List<int>();
                                        isCustomPages = false;
                                        break;
                                    }
                                    else
                                    {
                                        pageList.Add(int.Parse(str));
                                    }
                                }
                                catch
                                {
                                    MessageWindow window = new MessageWindow("The custom pages value is invalid.", "Error", "OK", null, MessageWindow.MessageIcon.Error)
                                    {
                                        Owner = Window.GetWindow(this),
                                        Icon = Window.GetWindow(this).Icon
                                    };
                                    window.ShowDialog();

                                    pageList = new List<int>();
                                    isCustomPages = false;
                                    break;
                                }
                            }
                        }
                    }
					lastPageList = pageList;
					#endregion

					#region Label Quantity
					double labelQuantityFactor = 0;
                    if (quantityComboBox.SelectedIndex == 1)
                    {
                        labelQuantityFactor = 100;
                    }
                    else if (quantityComboBox.SelectedIndex == 2 && String.IsNullOrWhiteSpace(customQuantityTextBox.Text) == false)
                    {
						try
						{
							labelQuantityFactor = Convert.ToInt32(customQuantityTextBox.Text);
						}
						catch (Exception)
						{

						}
                    }
					#endregion

					#region Color
					PrintSettings.PageColor color;
                    if (colorComboBox.SelectedItem != null && (colorComboBox.SelectedItem as ComboBoxItem).Content.ToString() == PrintSettings.PageColor.Grayscale.ToString())
                    {
                        ScrollViewer scrollViewer = (ScrollViewer)documentPreviewer.Template.FindName("PART_ContentHost", documentPreviewer);
                        scrollViewer.Effect = new PreviewHelper.GrayscaleEffect();

                        color = PrintSettings.PageColor.Grayscale;
                    }
                    else if (colorComboBox.SelectedItem != null && (colorComboBox.SelectedItem as ComboBoxItem).Content.ToString() == PrintSettings.PageColor.Monochrome.ToString())
                    {
                        ScrollViewer scrollViewer = (ScrollViewer)documentPreviewer.Template.FindName("PART_ContentHost", documentPreviewer);
                        scrollViewer.Effect = new PreviewHelper.GrayscaleEffect();

                        color = PrintSettings.PageColor.Monochrome;
                    }
                    else
                    {
                        ScrollViewer scrollViewer = (ScrollViewer)documentPreviewer.Template.FindName("PART_ContentHost", documentPreviewer);
                        scrollViewer.Effect = null;

                        color = PrintSettings.PageColor.Color;
                    }
					#endregion

					#region Margin
					double margin;
                    // default
                    if (marginComboBox.SelectedIndex == 0)
                    {
                        margin = _pageMargin;
                    }
                    else if (marginComboBox.SelectedIndex == 1)
                    {// none
                        margin = 0;
                    }
                    else if (marginComboBox.SelectedIndex == 2)
                    {// minimum
                        try
                        {
                            //todo check
                            margin = 0;
                            //PageImageableArea imageableArea = _printer.GetPrintCapabilities(new PrintTicket() { PageMediaSize = size }).PageImageableArea;
                            //margin = Math.Max(imageableArea.OriginHeight, imageableArea.OriginWidth);
                        }
                        catch
                        {
                            margin = 0;
                        }
                    }
                    else//3: Custom
                    {
                        try
                        {
                            margin = int.Parse(customMarginNumberPicker.Text);
                        }
                        catch
                        {
                            margin = 0;
                        }
                    }
					#endregion

					#region Scale
					double scale;
                    if (scaleComboBox.SelectedIndex == 0) // auto fit
                    {
                        scale = double.NaN;
                    }
                    else if (scaleComboBox.SelectedIndex == 7) // custom
                    {
                        try
                        {
                            scale = int.Parse(customZoomNumberPicker.Text);
                        }
                        catch
                        {
                            scale = 100;
                        }
                    }
                    else
                    {
                        scale = _zoomList[scaleComboBox.SelectedIndex - 1];
                    }
					#endregion

					#region Page Content
					List<PageContent> pageContentList;

                    if (_getDocumentWhenReloadDocumentMethod != null)
                    {
                        pageContentList = _getDocumentWhenReloadDocumentMethod(new PrintDialog.DocumentInfo()
                        {
                            Color = color,
                            Margin = margin,
                            Orientation = (PrintSettings.PageOrientation)orientationComboBox.SelectedIndex,
                            PageOrder = (PrintSettings.PageOrder)pageOrderComboBox.SelectedIndex,
                            Pages = pageList.ToArray(),
                            PagesPerSheet = int.Parse((pagesPerSheetComboBox.SelectedItem as ComboBoxItem).Content.ToString()),
                            LabelQuantityFactor = labelQuantityFactor,
                            Scale = scale,
                            Size = new Size(size.Width, size.Height),
                        });
                    }
                    else
                    {
                        pageContentList = _orginalPagesContentList;
                    }
					#endregion

					#region Load Pages
					int pageCount = 1;

                    foreach (PageContent orginalPage in pageContentList)
                    {
                        if ((isCustomPages == true && pageList.Count >= 1 && pageList.Contains(pageCount) == true) || isCustomPages == false)
                        {
                            FixedPage fixedPage = XamlReader.Parse(XamlWriter.Save(orginalPage.Child)) as FixedPage;

                            fixedPage.Width = doc.DocumentPaginator.PageSize.Width;
                            fixedPage.Height = doc.DocumentPaginator.PageSize.Height;
                            fixedPage.RenderTransformOrigin = new Point(0, 0);

                            if (pagesPerSheetComboBox.SelectedIndex == 0)
                            {
                                if (scaleComboBox.SelectedIndex == 0)
                                {
                                    if (_orginalPageSize.Height * (fixedPage.Width / _orginalPageSize.Width) <= fixedPage.Height)
                                    {
                                        fixedPage.RenderTransform = new ScaleTransform(fixedPage.Width / _orginalPageSize.Width, fixedPage.Width / _orginalPageSize.Width);
                                    }
                                    else
                                    {
                                        fixedPage.RenderTransform = new ScaleTransform(fixedPage.Height / _orginalPageSize.Height, fixedPage.Height / _orginalPageSize.Height);
                                    }
                                }
                                else
                                {
                                    fixedPage.RenderTransform = new ScaleTransform(scale / 100.0, scale / 100.0);
                                }

                                double finalMargin = 0 - _pageMargin + margin;

                                foreach (UIElement element in GetChildren(fixedPage))
                                {
									FixedPage.SetLeft(element, FixedPage.GetLeft(element) + finalMargin);
									FixedPage.SetTop(element, FixedPage.GetTop(element) + finalMargin);
									FixedPage.SetRight(element, FixedPage.GetRight(element) + finalMargin);
									FixedPage.SetBottom(element, FixedPage.GetBottom(element) + finalMargin);
								}
                            }

                            doc.Pages.Add(new PageContent { Child = fixedPage });

                            fixedPage.UpdateLayout();
                            DoEvents();
                        }

                        pageCount++;
                    }

                    // Pages Per Sheet
                    if (pagesPerSheetComboBox.SelectedIndex != 0)
                    {
                        PreviewHelper.MultiPagesPerSheetHelper multiPagesPerSheetHelper = new PreviewHelper.MultiPagesPerSheetHelper(int.Parse((pagesPerSheetComboBox.SelectedItem as ComboBoxItem).Content.ToString()), doc, _orginalPageSize, (PreviewHelper.DocumentOrientation)orientationComboBox.SelectedIndex, (PreviewHelper.PageOrder)pageOrderComboBox.SelectedIndex);
                        fixedDocument = multiPagesPerSheetHelper.GetMultiPagesPerSheetDocument(scale);
                    }
                    else
                    {
                        fixedDocument = doc;
                    }
					#endregion

                    LoadPreview(fixedDocument);
                }
                catch (Exception ex)
                {

                }
                finally
                {
                    loadingGrid.Visibility = Visibility.Collapsed;
                    DoEvents();
                }
            }
        }

        #endregion

        #region Define PrintDocument() method to print the document with user settings

        private void PrintDocument()
        {
            if (printerComboBox.SelectedItem == null)
            {
                return;
            }

            //ReloadDocument();

            _printer = _localDefaultPrintServer.GetPrintQueue((printerComboBox.SelectedItem as ComboBoxItem).Tag.ToString());
            SizeList = GetSizeList(_printer);

			PrintTicket printTicket = _systemPrintDialog.PrintTicket;

			#region get printTicket
			//PrintQueue printQueue = null;

			//LocalPrintServer localPrintServer = new LocalPrintServer();
			//// Retrieving collection of local printer on user machine
			//PrintQueueCollection localPrinterCollection =
			//	localPrintServer.GetPrintQueues();
			//System.Collections.IEnumerator localPrinterEnumerator =
			//	localPrinterCollection.GetEnumerator();

			//if (localPrinterEnumerator.MoveNext())
			//{
			//	// Get PrintQueue from first available printer
			//	printQueue = (PrintQueue)localPrinterEnumerator.Current;
			//}
			//// Get default PrintTicket from printer
			//PrintTicket printTicket = printQueue.DefaultPrintTicket;

			//// Modify PrintTicket
			//         PrintCapabilities printCapabilities = printQueue.GetPrintCapabilities();
			//if (printCapabilities.CollationCapability.Contains(Collation.Collated))
			//{
			//	printTicket.Collation = Collation.Collated;
			//}

			//if (printCapabilities.DuplexingCapability.Contains(
			//		Duplexing.TwoSidedLongEdge))
			//{
			//	printTicket.Duplexing = Duplexing.TwoSidedLongEdge;
			//}

			//if (printCapabilities.StaplingCapability.Contains(Stapling.StapleDualLeft))
			//{
			//	printTicket.Stapling = Stapling.StapleDualLeft;
			//}
			#endregion

			#region Setting
			// copy
			printTicket.CopyCount = int.Parse(copiesNumberPicker.Text);
			if (collateCheckBox.IsChecked == true)
			{
				printTicket.Collation = Collation.Collated;
			}
			else
			{
				printTicket.Collation = Collation.Uncollated;
			}
			// orientation
			if (orientationComboBox.SelectedIndex == 0)
			{
				printTicket.PageOrientation = PageOrientation.Portrait;
			}
			else
			{
				printTicket.PageOrientation = PageOrientation.Landscape;
			}
			// page media type
			if (_printer.GetPrintCapabilities().PageMediaTypeCapability.Count > 0)
			{
				printTicket.PageMediaType = _printer.GetPrintCapabilities().PageMediaTypeCapability[typeComboBox.SelectedIndex];
			}
			// color
			if (_printer.GetPrintCapabilities().OutputColorCapability.Count > 0)
			{
				printTicket.OutputColor = _printer.GetPrintCapabilities().OutputColorCapability[colorComboBox.SelectedIndex];
			}
			// quality
			if (_printer.GetPrintCapabilities().OutputQualityCapability.Count > 0)
			{
				printTicket.OutputQuality = _printer.GetPrintCapabilities().OutputQualityCapability[qualityComboBox.SelectedIndex];
			}
			// input bin
			if (_printer.GetPrintCapabilities().InputBinCapability.Count > 0)
			{
				printTicket.InputBin = _printer.GetPrintCapabilities().InputBinCapability[sourceComboBox.SelectedIndex];
			}
			// twosided
			if (twoSidedCheckBox.IsChecked == true)
			{
				if (twoSidedTypeComboBox.SelectedIndex == 0)
				{
					printTicket.Duplexing = Duplexing.TwoSidedLongEdge;
				}
				else
				{
					printTicket.Duplexing = Duplexing.TwoSidedShortEdge;
				}
			}
			else
			{
				printTicket.Duplexing = Duplexing.OneSided;
			}

			#endregion
			printTicket.PageScalingFactor = 100;
			printTicket.PagesPerSheet = 1;
			printTicket.PagesPerSheetDirection = PagesPerSheetDirection.RightBottom;

			// size
			if (SizeList.Count > 0)
			{
                printTicket = ModifyPrintTicket(printTicket, SizeList[sizeComboBox.SelectedIndex], true);
                //todo add new print ticket
                //printTicket.PageMediaSize = _printer.GetPrintCapabilities().PageMediaSizeCapability[sizeComboBox.SelectedIndex];
			}

            //todo del
			//Double printableWidth = printTicket.PageMediaSize.Width.Value;
			//Double printableHeight = printTicket.PageMediaSize.Height.Value;
   //         Double xMargin = 0;
   //         Double yMargin = 0;

			//Double xScale = (printableWidth - xMargin * 2) / printableWidth;
			//Double yScale = (printableHeight - yMargin * 2) / printableHeight;


			//this.LayoutTransform = new MatrixTransform(xScale, 0, 0, yScale, xMargin, yMargin);


			_systemPrintDialog.PrintQueue = _printer;
            _systemPrintDialog.PrintDocument(fixedDocument.DocumentPaginator, _documentName);
        }

        #endregion

        //Events

        #region The window loaded event to load printers, settings and document

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            DoEvents();

            if (_allowPagesOption == false)
            {
                pagesTextBlock.Visibility = Visibility.Collapsed;
                pagesComboBox.Visibility = Visibility.Collapsed;
            }

            if (_allowScaleOption == false)
            {
                scaleTextBlock.Visibility = Visibility.Collapsed;
                scaleComboBox.Visibility = Visibility.Collapsed;
                scaleComboBox.SelectedIndex = 4;
            }

            if (_allowTwoSidedOption == false)
            {
                twoSidedTextBlock.Visibility = Visibility.Collapsed;
                twoSidedCheckBox.Visibility = Visibility.Collapsed;
            }

            if (_allowPagesPerSheetOption == false)
            {
                pagesPerSheetTextBlock.Visibility = Visibility.Collapsed;
                pagesPerSheetComboBox.Visibility = Visibility.Collapsed;
            }
            else
            {
                if (_allowPageOrderOption == false)
                {
                    pageOrderTextBlock.Visibility = Visibility.Collapsed;
                    pageOrderComboBox.Visibility = Visibility.Collapsed;
                }
            }

            if (_allowMoreSettingsExpander == false)
            {
                moreSettingsExpander.MinHeight = 0;
                moreSettingsExpander.IsExpanded = true;
            }

            if (_allowPrinterPreferencesButton == false)
            {
                printerPreferencesBtn.Visibility = Visibility.Collapsed;
            }

            this.UpdateLayout();
            DoEvents();

            LoadPrinters();
            LoadPrinterSettings();

            PrintQueue printer = _localDefaultPrintServer.GetPrintQueue((printerComboBox.SelectedItem as ComboBoxItem).Tag.ToString());

            if (_defaultSettings.Layout == PrintSettings.PageOrientation.Portrait)
            {
                orientationComboBox.SelectedIndex = 0;
            }
            else
            {
                orientationComboBox.SelectedIndex = 1;
            }

            if (_allowTwoSidedOption == true)
            {
                if (_defaultSettings.UsePrinterDefaultSettings == false)
                {
                    if (_defaultSettings.TwoSided == PrintSettings.TwoSided.OneSided)
                    {
                        twoSidedCheckBox.IsChecked = false;
                    }
                    else
                    {
                        twoSidedCheckBox.IsChecked = true;

                        if (_defaultSettings.TwoSided == PrintSettings.TwoSided.TwoSidedLongEdge)
                        {
                            twoSidedTypeComboBox.SelectedIndex = 0;
                        }
                        else
                        {
                            twoSidedTypeComboBox.SelectedIndex = 1;
                        }
                    }
                }
                else
                {
                    if (printer.DefaultPrintTicket.Duplexing.HasValue)
                    {
                        if (printer.DefaultPrintTicket.Duplexing.Value == Duplexing.TwoSidedLongEdge)
                        {
                            twoSidedCheckBox.IsChecked = true;
                            twoSidedTypeComboBox.SelectedIndex = 0;
                        }
                        else if (printer.DefaultPrintTicket.Duplexing.Value == Duplexing.TwoSidedShortEdge)
                        {
                            twoSidedCheckBox.IsChecked = true;
                            twoSidedTypeComboBox.SelectedIndex = 1;
                        }
                        else
                        {
                            twoSidedCheckBox.IsChecked = false;
                        }
                    }
                    else
                    {
                        twoSidedCheckBox.IsChecked = false;
                    }
                }
            }

            pagesPerSheetComboBox.SelectedIndex = PreviewHelper.MultiPagesPerSheetHelper.GetPagePerSheetCountIndex(_defaultSettings.PagesPerSheet);
            pageOrderComboBox.SelectedIndex = (int)_defaultSettings.PageOrder;

            customMarginNumberPicker.MaxValue = (int)Math.Min(_orginalPageSize.Width / 2, _orginalPageSize.Height / 2) - 15;
            customMarginNumberPicker.Text = _pageMargin.ToString();

            DoEvents();

            ReloadDocument();
            documentPreviewer.FitToWidth();

            loadingGrid.Visibility = Visibility.Collapsed;
            DoEvents();

            documentPreviewer.FitToWidth();
            DoEvents();

            isLoaded = true;

            ((TextBlock)documentPreviewer.Template.FindName("currentPageTextBlock", documentPreviewer)).Text = "Page " + GetCurrentPage().ToString() + " / " + documentPreviewer.PageCount.ToString();
            printButton.Focus();
        }

        #endregion

        #region The window closed event to release properties

        private void Window_Closed(object sender, EventArgs e)
        {
            PackageStore.RemovePackage(_xpsUrl);
            if (package != null)
            {
                package.Close();
            }

            _localDefaultPrintServer.Dispose();
        }

        #endregion

        #region The print button click event to print document

        private void PrintButton_Click(object sender, RoutedEventArgs e)
        {
            cancelButton.IsEnabled = false;
            printButton.IsEnabled = false;
            DoEvents();

            PrintDocument();

            ReturnValue = true;
            Window.GetWindow(this).Close();
        }

        #endregion

        #region The cancel button click event to close the window

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            cancelButton.IsEnabled = false;
            printButton.IsEnabled = false;
            DoEvents();

            ReturnValue = false;
            Window.GetWindow(this).Close();
        }

        #endregion

        #region The printer preferences button click event to open printer preferences dialog

        private void PrinterPreferencesButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Drawing.Printing.PrinterSettings settings = new System.Drawing.Printing.PrinterSettings
                {
                    PrinterName = (printerComboBox.SelectedItem as ComboBoxItem).Tag.ToString()
                };
                OpenPrinterPropertiesDialog(settings);
            }
            catch
            {
                return;
            }
        }

        #endregion

        #region The printer combo box drop down opened event to refresh printer list

        private void EquipmentComboBox_DropDownOpened(object sender, EventArgs e)
        {
            isRefresh = false;

            _localDefaultPrintServer.Refresh();

            string equipmentComboBoxSelectedItem = (printerComboBox.SelectedItem as ComboBoxItem).Tag.ToString();
            int equipmentComboBoxSelectedIndex = -1;
            int defaultEquipmentComboBoxSelectedIndex = -1;
            printerComboBox.Items.Clear();

            if (_allowAddNewPrinerComboBoxItem == true)
            {
                Grid itemMainGrid = new Grid();
                itemMainGrid.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(55) });
                itemMainGrid.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(10) });
                itemMainGrid.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(1, GridUnitType.Star) });

                TextBlock textInfoBlock = new TextBlock()
                {
                    Text = "Add New Printer",
                    FontSize = 14,
                    VerticalAlignment = VerticalAlignment.Center
                };

                Image itemIcon = new Image()
                {
                    Width = 55,
                    Height = 55,
                    Source = new System.Windows.Media.Imaging.BitmapImage(new Uri("/PrintDialog;component/Resources/AddPrinter.png", UriKind.Relative)),
                    Stretch = Stretch.Fill
                };

                itemIcon.SetValue(Grid.ColumnProperty, 0);
                textInfoBlock.SetValue(Grid.ColumnProperty, 2);

                itemMainGrid.Children.Add(itemIcon);
                itemMainGrid.Children.Add(textInfoBlock);

                ComboBoxItem item = new ComboBoxItem
                {
                    Content = itemMainGrid,
                    Height = 55,
                    ToolTip = "Add New Printer"
                };

                printerComboBox.Items.Insert(0, item);
            }

            foreach (PrintQueue printer in _localDefaultPrintServer.GetPrintQueues())
            {
                printer.Refresh();

                string status = PrinterHelper.PrinterHelper.GetPrinterStatusInfo(_localDefaultPrintServer.GetPrintQueue(printer.FullName));
                string loction = printer.Location;
                string comment = printer.Comment;

                if (String.IsNullOrWhiteSpace(loction) == true)
                {
                    loction = "Unknown";
                }
                if (String.IsNullOrWhiteSpace(comment) == true)
                {
                    comment = "Unknown";
                }

                Grid itemMainGrid = new Grid();
                itemMainGrid.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(55) });
                itemMainGrid.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(10) });
                itemMainGrid.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(1, GridUnitType.Star) });

                StackPanel textInfoPanel = new StackPanel()
                {
                    Orientation = Orientation.Vertical,
                    VerticalAlignment = VerticalAlignment.Center
                };
                textInfoPanel.Children.Add(new TextBlock() { Text = printer.FullName, FontSize = 14 });
                textInfoPanel.Children.Add(new TextBlock() { Text = status, Margin = new Thickness(0, 7, 0, 0) });

                Image printerIcon = new Image()
                {
                    Width = 55,
                    Height = 55,
                    Source = PrinterInfoHelper.PrinterIconHelper.GetPrinterIcon(printer, _localDefaultPrintServer),
                    Stretch = Stretch.Fill
                };

                if (printer.IsOffline == true)
                {
                    printerIcon.Opacity = 0.5;
                }

                printerIcon.SetValue(Grid.ColumnProperty, 0);
                textInfoPanel.SetValue(Grid.ColumnProperty, 2);

                itemMainGrid.Children.Add(printerIcon);
                itemMainGrid.Children.Add(textInfoPanel);

                ComboBoxItem item = new ComboBoxItem
                {
                    Content = itemMainGrid,
                    Height = 55,
                    ToolTip = "Name: " + printer.FullName + "\nStatus: " + status + "\nDocument: " + printer.NumberOfJobs + "\nLoction: " + loction + "\nComment: " + comment,
                    Tag = printer.FullName
                };

                printerComboBox.Items.Insert(0, item);

                if (equipmentComboBoxSelectedItem == printer.FullName)
                {
                    equipmentComboBoxSelectedIndex = printerComboBox.Items.Count;
                }
                if (LocalPrintServer.GetDefaultPrintQueue().FullName == printer.FullName)
                {
                    defaultEquipmentComboBoxSelectedIndex = printerComboBox.Items.Count;
                }
            }

            if (equipmentComboBoxSelectedIndex == -1)
            {
                if (defaultEquipmentComboBoxSelectedIndex == -1)
                {
                    equipmentComboBoxSelectedIndex = 0;
                }
                else
                {
                    equipmentComboBoxSelectedIndex = printerComboBox.Items.Count - defaultEquipmentComboBoxSelectedIndex;
                }
            }
            else
            {
                equipmentComboBoxSelectedIndex = printerComboBox.Items.Count - equipmentComboBoxSelectedIndex;
            }

            lastPrinterIndex = equipmentComboBoxSelectedIndex;
            printerComboBox.SelectedIndex = equipmentComboBoxSelectedIndex;

            printerComboBox.Tag = PrinterInfoHelper.PrinterIconHelper.GetPrinterIcon(_localDefaultPrintServer.GetPrintQueue((printerComboBox.SelectedItem as ComboBoxItem).Tag.ToString()), _localDefaultPrintServer, true);
            printerComboBox.Text = (printerComboBox.SelectedItem as ComboBoxItem).Tag.ToString();

            isRefresh = true;
        }

        #endregion

        //The printer combo box selection changed event

        private async void EquipmentComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (printerComboBox.SelectedItem == null)
            {
                return;
            }

            if (isLoaded == true && isRefresh == true)
            {
                printerComboBox.IsDropDownOpen = false;

                if (printerComboBox.SelectedIndex == printerComboBox.Items.Count - 1 && (printerComboBox.SelectedItem as ComboBoxItem).Tag == null)
                {
                    printerComboBox.SelectedIndex = lastPrinterIndex;

                    printerComboBox.Tag = PrinterInfoHelper.PrinterIconHelper.GetPrinterIcon(_localDefaultPrintServer.GetPrintQueue((printerComboBox.SelectedItem as ComboBoxItem).Tag.ToString()), _localDefaultPrintServer, true);
                    printerComboBox.Text = (printerComboBox.SelectedItem as ComboBoxItem).Tag.ToString();

                    // todo add uwp: https://www.nuget.org/packages/WindowsContracts.Net.Foundation.UniversalApiContract/5.19041.10/license
                    Windows.System.LauncherOptions option = new Windows.System.LauncherOptions()
                    {
                        TreatAsUntrusted = false
                    };
                    bool result = await Windows.System.Launcher.LaunchUriAsync(new Uri("ms-settings:printers"), option);

                    if (result == false)
                    {
                        System.Diagnostics.Process.Start("explorer.exe", "shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}");
                    }
                }
                else
                {
                    if (lastPrinterIndex != printerComboBox.SelectedIndex)
                    {
                        printerComboBox.Tag = PrinterInfoHelper.PrinterIconHelper.GetPrinterIcon(_localDefaultPrintServer.GetPrintQueue((printerComboBox.SelectedItem as ComboBoxItem).Tag.ToString()), _localDefaultPrintServer, true);
                        printerComboBox.Text = (printerComboBox.SelectedItem as ComboBoxItem).Tag.ToString();

                        lastPrinterIndex = printerComboBox.SelectedIndex;
                        LoadPrinterSettings(false);

                        printerComboBox.Tag = PrinterInfoHelper.PrinterIconHelper.GetPrinterIcon(_localDefaultPrintServer.GetPrintQueue((printerComboBox.SelectedItem as ComboBoxItem).Tag.ToString()), _localDefaultPrintServer, true);
                        printerComboBox.Text = (printerComboBox.SelectedItem as ComboBoxItem).Tag.ToString();
                    }
                    else
                    {
                        printerComboBox.Tag = PrinterInfoHelper.PrinterIconHelper.GetPrinterIcon(_localDefaultPrintServer.GetPrintQueue((printerComboBox.SelectedItem as ComboBoxItem).Tag.ToString()), _localDefaultPrintServer, true);
                        printerComboBox.Text = (printerComboBox.SelectedItem as ComboBoxItem).Tag.ToString();
                    }
                }
            }
        }

        #region The setting combo boxes selection changed event to reload document

        private void SettingComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (isLoaded == true)
            {
                if (sender is ComboBox comboBox)
                {
                    comboBox.IsDropDownOpen = false;
                }
                //scale
                if (scaleComboBox.SelectedIndex == 7)
                {
                    customZoomNumberPicker.Visibility = Visibility.Visible;
                }
                else
                {
                    customZoomNumberPicker.Visibility = Visibility.Collapsed;
                }
                //pages
                if (pagesComboBox.SelectedIndex == 2)
                {
                    customPagesTextBox.Visibility = Visibility.Visible;
                }
                else
                {
                    customPagesTextBox.Visibility = Visibility.Collapsed;
                }

                //label quantity
                if (quantityComboBox.SelectedIndex == 2)
                {
                    customQuantityTextBox.Visibility = Visibility.Visible;
                }
                else
                {
                    customQuantityTextBox.Visibility = Visibility.Collapsed;
                }
                //margin
                if (marginComboBox.SelectedIndex == 3)
                {
                    customMarginNumberPicker.Visibility = Visibility.Visible;
                }
                else
                {
                    customMarginNumberPicker.Visibility = Visibility.Collapsed;
                }

                if (pagesPerSheetComboBox.SelectedIndex == 0)
                {
                    marginComboBox.IsEnabled = true;
                    customMarginNumberPicker.IsEnabled = true;
                    (pagesComboBox.Items[1] as ComboBoxItem).Visibility = Visibility.Visible;
                }
                else
                {
                    marginComboBox.IsEnabled = false;
                    customMarginNumberPicker.IsEnabled = false;

                    if (pagesComboBox.SelectedIndex != 1)
                    {
                        (pagesComboBox.Items[1] as ComboBoxItem).Visibility = Visibility.Collapsed;
                    }
                }

                ReloadDocument();
            }
        }

        #endregion

        #region The document viewer context menu opening event to cancel it

        private void DocumentPreviewer_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            e.Handled = true;
        }

        #endregion

        #region The document viewer manipulation boundary feedback event to cancel it

        private void DocumentPreviewer_ManipulationBoundaryFeedback(object sender, System.Windows.Input.ManipulationBoundaryFeedbackEventArgs e)
        {
            e.Handled = true;
        }

        #endregion

        #region The document viewer scroll changed event to refresh the current page text block content

        private void DocumentPreviewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            ((TextBlock)documentPreviewer.Template.FindName("currentPageTextBlock", documentPreviewer)).Text = "Page " + GetCurrentPage().ToString() + " / " + documentPreviewer.PageCount.ToString();
        }

        #endregion

        #region The document viewer navigate buttons click events to navigate

        private void FirstPageBtn_Click(object sender, RoutedEventArgs e)
        {
            documentPreviewer.FirstPage();
        }

        private void PreviousPageBtn_Click(object sender, RoutedEventArgs e)
        {
            documentPreviewer.PreviousPage();
        }

        private void NextPageBtn_Click(object sender, RoutedEventArgs e)
        {
            documentPreviewer.NextPage();
        }

        private void LastPageBtn_Click(object sender, RoutedEventArgs e)
        {
            documentPreviewer.LastPage();
        }

        #endregion

        #region The document viewer actual size button click event to cancel 2 pages across mode

        private void ActualSizeBtn_Click(object sender, RoutedEventArgs e)
        {
            documentPreviewer.FitToMaxPagesAcross(1);
            documentPreviewer.Zoom = 100.0;
        }

        #endregion

        #region The copies count number picker input completed event to show collate check box or not

        private void CopiesNumberPicker_InputCompleted(object sender, RoutedEventArgs e)
        {
            if (isLoaded == true)
            {
                if (int.Parse(copiesNumberPicker.Text) > 1)
                {
                    collateCheckBox.Visibility = Visibility.Visible;
                }
                else
                {
                    collateCheckBox.Visibility = Visibility.Collapsed;
                }
            }
        }

        #endregion

        #region The custom zoom number picker input completed event to reload document

        private void CustomZoomNumberPicker_InputComplated(object sender, RoutedEventArgs e)
        {
            if (isLoaded == true)
            {
                ReloadDocument();
            }
        }

        #endregion

        #region The custom margin number picker input completed event to reload document

        private void CustomMarginNumberPicker_InputCompleted(object sender, RoutedEventArgs e)
        {
            if (isLoaded == true)
            {
                ReloadDocument();
            }
        }

        #endregion

        #region The custom pages text box lost focus event to reload document

        private void CustomPagesTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (isLoaded == true)
            {
                ReloadDocument();
            }
        }

        private void CustomQuantityTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (isLoaded == true)
            {
                ReloadDocument();
            }
        }

        #endregion

        protected static string GetMD5Hash(byte[] data)
        {
            MD5 md5 = MD5.Create();
            md5.Initialize();
            md5.TransformFinalBlock(data, 0, data.Length);
            return new String(md5.Hash.SelectMany((b) => b.ToString("X2").ToCharArray()).ToArray());
        }

        public class PageSize
        {
            public string Name { get; set; }
			public int NameID { get; set; }
			public double Width { get; set; }
			public double Height { get; set; }
            // original value
            public string nameStr { get; set; }
            public string widthStr { get; set; }
            public string heightStr { get; set; }
		}

        /// <summary>
        /// Modifes a print ticket xml after updating a feature value.
        /// 
        /// Sample usage:
        /// Get Dictionary with Inputbins by calling the other method
        /// and get "value" for the desired inputbin you'd like to use...
        /// ...
        /// desiredTray is then something like "NS0000:SurpriseOption7" for example.
        /// defaultPrintTicket is the (Default)PrintTicket you want to modify from the PrintQueue for example
        /// </summary>
        /// <param name="ticket"></param>
        /// <param name="featureName"></param>
        /// <param name="newValue"></param>
        /// <param name="printQueueName">Optional - If provided, a file is created with the print ticket xml. Useful for debugging.</param>
        /// <param name="folder">Optional - If provided, the path for a file is created with the print ticket xml. Defaults to c:\. Useful for debugging.</param>
        /// <param name="displayMessage">Optional - True to display a dialog with changes. Defaults to false. Useful for debugging.</param>
        /// <returns></returns>
        /// PrintTicket myPrintTicket = WpfPrinterUtils.ModifyPrintTicket(defaultPrintTicket, "psk:JobInputBin", desiredTray);
        /// 
        public static PrintTicket ModifyPrintTicket(PrintTicket ticket, PageSize size, bool showXML = false)
        {
            bool displayMessage = false;
            string xmlFolder = System.Windows.Forms.Application.StartupPath + @"\temp";

            if (ticket == null)
                throw new ArgumentNullException("ticket");

            string newName = size.nameStr;

            // Read Xml of the PrintTicket xml.
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(ticket.GetXmlStream());

            // Create NamespaceManager and add PrintSchemaFrameWork-Namespace hinzufugen (should be on DocumentElement of the PrintTicket).
            // Prefix: psf NameSpace: xmlDoc.DocumentElement.NamespaceURI = "http://schemas.microsoft.com/windows/2003/08/printing/printschemaframework"
            XmlNamespaceManager manager = new XmlNamespaceManager(xmlDoc.NameTable);
            manager.AddNamespace(xmlDoc.DocumentElement.Prefix, xmlDoc.DocumentElement.NamespaceURI);

            // Search node with desired feature we're looking for and set newValue for it
            string xpath = string.Format("//psf:Feature[@name='{0}']/psf:Option", "psk:PageMediaSize");
            XmlNode node = xmlDoc.SelectSingleNode(xpath, manager);
            if (node != null)
            {
                if (node.Attributes["name"].Value != newName)
                {
                    if (displayMessage)
                    {
                        MessageBox.Show(string.Format("OldValue: {0}, NewValue: {1}", node.Attributes["name"].Value, newName), "Size Name");
                    }
                    node.Attributes["name"].Value = newName;
                    node.ChildNodes[0].InnerText = size.widthStr;
                    node.ChildNodes[1].InnerText = size.heightStr;
                    //todo add
                }
            }

            // Create a new PrintTicket out of the XML.
            PrintTicket newPrintTicket = null;
            using (MemoryStream stream = new MemoryStream())
            {
                xmlDoc.Save(stream);
                stream.Position = 0;
                newPrintTicket = new PrintTicket(stream);
            }

            // For testing purpose save the print ticket to a file.
            if (showXML)
            {
                if (string.IsNullOrWhiteSpace(xmlFolder))
                {
                    xmlFolder = "C:\\";
                }
                string fileFullPath = xmlFolder + @"\printTicket.xml";
                if (File.Exists(fileFullPath))
                {
                    File.Delete(fileFullPath);
                }
                if (!Directory.Exists(Path.GetDirectoryName(fileFullPath)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(fileFullPath));
                }
                using (FileStream stream = new FileStream(fileFullPath, FileMode.CreateNew, FileAccess.ReadWrite))
                {
                    newPrintTicket.GetXmlStream().WriteTo(stream);
                }
            }

            if (newPrintTicket.PageMediaSize.Width == null)
            {
				PageMediaSize newSize;

                if(newPrintTicket.PageMediaSize.PageMediaSizeName != null)
                    newSize = new PageMediaSize((PageMediaSizeName)newPrintTicket.PageMediaSize.PageMediaSizeName, size.Width, size.Height);
                else 
                    newSize = new PageMediaSize(size.Width, size.Height);
				
                newPrintTicket.PageMediaSize = newSize;
            }

            return newPrintTicket;
        }
        // End
    }
}