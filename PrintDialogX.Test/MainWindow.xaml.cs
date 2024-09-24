using System.Data;
using System.Drawing;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;

// print
using ZXing;
using ZXing.Common;
using Size = System.Windows.Size;
using MessageBox = System.Windows.MessageBox;
using Brushes = System.Windows.Media.Brushes;
using Point = System.Windows.Point;
using Image = System.Windows.Controls.Image;
using QRCoder;
using System.Windows.Markup;
using PrintDialogX.Common;

namespace PrintDialogX.Test;
/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    #region app config
    private static string g_BaseFolder = System.Windows.Forms.Application.StartupPath;
    private static string g_TempFolder = g_BaseFolder + @"\temp\";
    private static string g_ConfigFolder = g_BaseFolder + "Config\\";
    private static string g_NameSpace = "PrintDialogX";
    #endregion


    public MainWindow()
    {
        InitializeComponent();
    }

    private void OutputReport(object sender, RoutedEventArgs e)
    {
        // demo

        // size
        _labelWidth = 320;
        _labelHeight = 240;
        _labelNameInProject = "DemoPartLabel";

        // test data

        var partDtos = new List<PartLabelDto>()
        {
            new PartLabelDto(){PartRef = "Part_1", PartName = "part1 name", MatRef = "Mat01", Quantity = 1, Barcode = "part 1 barcode", QRCode = "part 1 qr code"},
            new PartLabelDto(){PartRef = "Part_2", PartName = "part2 name", MatRef = "Mat02", Quantity = 2, Barcode = "part 2 barcode", QRCode = "part 2 qr code"},
            new PartLabelDto(){PartRef = "Part_3", PartName = "part3 name", MatRef = "Mat03", Quantity = 10, Barcode = "part 3 barcode", QRCode = "part 3 qr code"},
        };

        // image
        foreach (var part in partDtos)
        {
            part.Image = Path.Combine(g_ConfigFolder, "test image", 
                part.PartRef + ".PNG");
        }

        _labelDatatable = partDtos.ToDataTable();

        ShowLabel();
    }

    #region show label

    public static PrintDialogX.PrintDialog.PrintDialog printDialog;
    private void ShowLabel()
    {

        printDialog = new PrintDialogX.PrintDialog.PrintDialog
        {
            Owner = this,
            Topmost = false,
            ShowInTaskbar = true,
            WindowStartupLocation = WindowStartupLocation.CenterOwner
        };

        if (printDialog.ShowDialog(true, LoadLabel) == true)
        {
            MessageBox.Show("Print OK");
        }
    }

    // set for each report

    public static DataTable _labelDatatable;
    public static Double _labelWidth, _labelHeight;
    public static string _labelNameInProject;
    // local file
    // string[] getFiles = Directory.GetFiles(templateFolder, "*.xaml");
    public static string _labelFilePath;

    // set for all report
    public static Double _labelMargin = 10;


    private static void LoadLabel()
    {
        FixedDocument fixedDocument = new FixedDocument();
        fixedDocument.DocumentPaginator.PageSize = new Size(_labelWidth + _labelMargin * 2, _labelHeight + _labelMargin * 2);

        foreach (DataRow dtRow in _labelDatatable.Rows)
        {
            FixedPage fixedPage = new FixedPage()
            {
                Width = fixedDocument.DocumentPaginator.PageSize.Width + _labelMargin * 2,
                Height = fixedDocument.DocumentPaginator.PageSize.Height + _labelMargin * 2
            };

            Grid grid = CreateLabel(dtRow, fixedDocument.DocumentPaginator.PageSize.Width, fixedDocument.DocumentPaginator.PageSize.Height, _labelMargin);

            fixedPage.Children.Add(grid);

            fixedDocument.Pages.Add(new PageContent() { Child = fixedPage });
        }

        //Setup PrintDialog's properties
        printDialog.Document = fixedDocument;
        printDialog.DocumentName = "标签打印";
        printDialog.DocumentMargin = _labelMargin;
        printDialog.DefaultSettings = PrintDialogX.PrintDialog.PrintDialogSettings.PrinterDefaultSettings();

        printDialog.AllowScaleOption = true; //缩放
        printDialog.AllowPagesOption = true; //页面范围 "All Pages", "Current Page", and "Custom Pages"
        printDialog.AllowTwoSidedOption = true; // 双面打印
        printDialog.AllowPagesPerSheetOption = true; //Allow pages per sheet option
        printDialog.AllowPageOrderOption = true; //Allow page order option
        printDialog.AllowAddNewPrinterButton = true; //Allow add new printer button in printer list
        printDialog.AllowMoreSettingsExpander = true; //Allow more settings expander
        printDialog.AllowPrinterPreferencesButton = true; //Allow printer preferences button

        printDialog.GetReloadPageMethod = GetReloadLabel; //Set the method that will use to recreate the document when print settings changed.

        printDialog.LoadingEnd();
    }

    public static Grid CreateLabel(DataRow dtRow, double width, double height, double margin)
    {
        try
        {
            Grid grid = new Grid()
            {
                Background = Brushes.White,
                Width = width,
                Height = height
            };

            // load xaml file in project via namespace
            var label = (UserControl)GetTypeInstance(g_NameSpace + "." + _labelNameInProject);
            
            // load local xaml file
            //var label = (UserControl)XamlReader.Load(new FileStream(g_LabelFilePath, FileMode.Open));
            
            label.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
            label.VerticalAlignment = System.Windows.VerticalAlignment.Center;

            foreach (var item in dtRow.Table.Columns)
            {
                string title = item.ToString();
                TextBlock textBlock = FindChildByName<TextBlock>((Grid)label.Content, title);
                if (textBlock != null)
                    textBlock.Text = dtRow[title].ToString();

                Image image = FindChildByName<Image>((Grid)label.Content, title);
                if (image != null)
                {

                    if (title == nameof(LabelDtoBase.Image) && File.Exists(dtRow[title].ToString()))
                        image.Source = ToBitmap(dtRow[title].ToString());
                    else if (title == nameof(LabelDtoBase.Barcode)) //Barcode string
                        image.Source = ToBarcodeImage(dtRow[title].ToString(), image.Height, image.Width);
                    else if (title == nameof(LabelDtoBase.QRCode)) //QRCode string
                        image.Source = ToQRCodeImage(dtRow[title].ToString(), image.Height, image.Width);

                    // todo check
                    if (title != nameof(LabelDtoBase.Image) && image.Source != null)
                        File.Delete(image.Source.ToString().Replace("file:///", ""));
                }
            }

            grid.Children.Add(label);

            FixedPage.SetTop(grid, margin); //Top margin
            FixedPage.SetLeft(grid, margin); //Left margin
            FixedPage.SetBottom(grid, margin); //Top margin
            FixedPage.SetRight(grid, margin); //Left margin

            return grid;
        }
        catch (Exception ex)
        {
            //ShowMessage(ex);
        }
        return new Grid();
    }

    public static List<PageContent> GetReloadLabel(PrintDialogX.PrintDialog.DocumentInfo documentInfo)
    {
        List<PageContent> pages = new List<PageContent>();

        foreach (DataRow dtRow in _labelDatatable.Rows)
        {
            FixedPage fixedPage;
            Grid grid;

            int quantity;
            int newQuantity = 1;
            double percent = 0;
            try
            {
                quantity = Convert.ToInt32(dtRow[nameof(LabelDtoBase.Quantity)].ToString());
                percent = (double)(documentInfo.LabelQuantityFactor / 100);
                if (percent > 1)
                    percent = 1;

                if (percent > 0)
                {
                    newQuantity = (int)(quantity * percent);

                    // 
                    if (newQuantity == 0)
                        newQuantity = 1;
                }
            }
            catch (Exception ex)
            {
            }

            for (int i = 0; i < newQuantity; i++)
            {
                fixedPage = new FixedPage()
                {
                    Width = _labelWidth + _labelMargin * 2,
                    Height = _labelHeight + _labelMargin * 2
                };
                grid = CreateLabel(dtRow, _labelWidth, _labelHeight, documentInfo.Margin.Value);
                fixedPage.Children.Add(grid);
                pages.Add(new PageContent() { Child = fixedPage });
            }
        }

        return pages;
    }
    #endregion

    #region UI Help function
    public static object GetTypeInstance(string strFullyQualifiedName)
    {
        Type t = Type.GetType(strFullyQualifiedName);
        return Activator.CreateInstance(t);
    }

    public static T FindChildByName<T>(DependencyObject parent, string name) where T : DependencyObject
    {
        for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            DependencyObject child = VisualTreeHelper.GetChild(parent, i);
            string a = child.GetValue(FrameworkElement.NameProperty) as string;
            if (a == name)
            {
                return child as T;
            }
            T val = FindChildByName<T>(child, name);
            if (val != null)
            {
                return val;
            }
        }
        return null;
    }
    #endregion

    #region Image, barcode, QR code
    public static BitmapImage ToBitmap(String path)
    {

        if (File.Exists(path))
        {
            BitmapImage bitImage = new BitmapImage();
            bitImage.BeginInit();
            bitImage.CacheOption = BitmapCacheOption.OnLoad;// Image can be deleted after rendering
            bitImage.UriSource = new Uri(path);
            bitImage.EndInit();
            return bitImage;
        }

        return null;
    }

    public static ImageSource ToBarcodeImage(string content, double height, double width)
    {//Barcode 不支持中文
        try
        {
            // todo check
            BarcodeWriter<Bitmap> writer = new BarcodeWriter<Bitmap>
            {
                Format = BarcodeFormat.CODE_39,
                Options = new EncodingOptions
                {
                    Height = (int)height,
                    Width = (int)width,
                    PureBarcode = true,
                    Margin = 10,
                }
            };

            var bitMap = writer.Write(content);
            string fileTemp = g_TempFolder + "barcode_" + GetTick() + ".bmp";
            FileHelper.CreateDirectory(g_TempFolder);
            bitMap.Save(fileTemp);

            Uri src = new Uri(fileTemp, UriKind.RelativeOrAbsolute);
            var bitImage = new BitmapImage();
            bitImage.BeginInit();
            bitImage.CacheOption = BitmapCacheOption.OnLoad;// Image can be deleted after rendering
            bitImage.UriSource = src;
            bitImage.EndInit();

            return bitImage;
        }
        catch (Exception ex)
        {
            return null;
        }
    }

    public static ImageSource ToQRCodeImage(string content, double height, double width)
    {
        try
        {
            // QrCoder
            QRCodeGenerator qrGenerator = new QRCodeGenerator();
            QRCodeData qrCodeData = qrGenerator.CreateQrCode(content, QRCodeGenerator.ECCLevel.Q);
            QRCoder.QRCode qrCode = new QRCoder.QRCode(qrCodeData);
            System.Drawing.Bitmap bitMap = qrCode.GetGraphic(20);

            // ZXing
            //BarcodeWriter writer = new BarcodeWriter
            //{
            //	Format = BarcodeFormat.QR_CODE,
            //	Options = new QrCodeEncodingOptions
            //	{
            //		DisableECI = true,//设置内容编码
            //		CharacterSet = "UTF-8",//设置内容编码
            //		Width = (int)width,
            //		Height = (int)height,
            //		Margin = 0
            //	}
            //};

            //var bitMap = writer.Write(content);
            string fileTemp = g_TempFolder + "qrCode_" + GetTick() + ".bmp";
            FileHelper.CreateDirectory(g_TempFolder);
            bitMap.Save(fileTemp);

            Uri src = new Uri(fileTemp, UriKind.RelativeOrAbsolute);
            var bitmapImage = new BitmapImage();
            bitmapImage.BeginInit();
            bitmapImage.CacheOption = BitmapCacheOption.OnLoad;// Image can be deleted after rendering
            bitmapImage.UriSource = src;
            bitmapImage.EndInit();

            return bitmapImage;
        }
        catch (Exception ex)
        {
            return null;
        }
    }
    #endregion

    #region common help function
    public static string GetTick()
    {
        return DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss_") + DateTime.Now.Ticks;
    }
    #endregion
}