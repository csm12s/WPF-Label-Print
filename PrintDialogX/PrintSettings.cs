using System.Printing;

namespace PrintDialogX.PrintSettings
{
    public enum PageOrientation
    {
        /// <summary>
        /// Standard orientation.
        /// </summary>
        Portrait,

        /// <summary>
        /// Content of the imageable area is rotated on the page 90 degrees counterclockwise from standard (portrait) orientation.
        /// </summary>
        Landscape
    }

    public enum PageColor
    {
        /// <summary>
        /// Output that prints in color.
        /// </summary>
        Color,

        /// <summary>
        /// Output that prints in a grayscale.
        /// </summary>
        Grayscale,

        /// <summary>
        /// Output that prints in a single color and with the same degree of intensity.
        /// </summary>
        Monochrome
    }

    public enum PageQuality
    {
        /// <summary>
        /// Automatically selects a quality type that is based on the contents of a print job.
        /// </summary>
        Automatic,

        /// <summary>
        /// Draft quality.
        /// </summary>
        Draft,

        /// <summary>
        /// Fax quality.
        /// </summary>
        Fax,

        /// <summary>
        /// Higher than normal quality.
        /// </summary>
        High,

        /// <summary>
        /// Normal quality.
        /// </summary>
        Normal,

        /// <summary>
        /// Photographic quality.
        /// </summary>
        Photographic,

        /// <summary>
        /// Text quality.
        /// </summary>
        Text
    }

    public enum PageType
    {
        /// <summary>
        /// The print device selects the media.
        /// </summary>
        AutoSelect,

        /// <summary>
        /// Archive-quality media.
        /// </summary>
        Archival,

        /// <summary>
        /// Specialty back-printing film.
        /// </summary>
        BackPrintFilm,

        /// <summary>
        /// Standard bond media.
        /// </summary>
        Bond,

        /// <summary>
        /// Standard card stock.
        /// </summary>
        CardStock,

        /// <summary>
        /// Continuous-feed media.
        /// </summary>
        Continuous,

        /// <summary>
        /// Standard envelope.
        /// </summary>
        EnvelopePlain,

        /// <summary>
        /// Window envelope.
        /// </summary>
        EnvelopeWindow,

        /// <summary>
        /// Fabric media.
        /// </summary>
        Fabric,

        /// <summary>
        /// Specialty high-resolution media.
        /// </summary>
        HighResolution,

        /// <summary>
        /// Label media.
        /// </summary>
        Label,

        /// <summary>
        /// Attached multipart forms.
        /// </summary>
        MultiLayerForm,

        /// <summary>
        /// Individual multipart forms.
        /// </summary>
        MultiPartForm,

        /// <summary>
        /// Standard photographic media.
        /// </summary>
        Photographic,

        /// <summary>
        /// Film photographic media.
        /// </summary>
        PhotographicFilm,

        /// <summary>
        /// Glossy photographic media.
        /// </summary>
        PhotographicGlossy,

        /// <summary>
        /// High-gloss photographic media.
        /// </summary>
        PhotographicHighGloss,

        /// <summary>
        /// Matte photographic media.
        /// </summary>
        PhotographicMatte,

        /// <summary>
        /// Satin photographic media.
        /// </summary>
        PhotographicSatin,

        /// <summary>
        /// Semi-gloss photographic media.
        /// </summary>
        PhotographicSemiGloss,

        /// <summary>
        /// Plain paper.
        /// </summary>
        Plain,

        /// <summary>
        /// Output to a display in continuous form.
        /// </summary>
        Screen,

        /// <summary>
        /// Output to a display in paged form.
        /// </summary>
        ScreenPaged,

        /// <summary>
        /// Specialty stationary.
        /// </summary>
        Stationery,

        /// <summary>
        /// Tab stock, not precut (single tabs).
        /// </summary>
        TabStockFull,

        /// <summary>
        /// Tab stock, precut (multiple tabs).
        /// </summary>
        TabStockPreCut,

        /// <summary>
        /// Transparent sheet.
        /// </summary>
        Transparency,

        /// <summary>
        /// Media that is used to transfer an image to a T-shirt.
        /// </summary>
        TShirtTransfer
    }

    public enum TwoSided
    {
        /// <summary>
        /// Output prints on only one side of each sheet.
        /// </summary>
        OneSided,

        /// <summary>
        /// Output prints on both sides of each sheet, which flips along the edge parallel to <see cref="PrintDocumentImageableArea.MediaSizeWidth"/>.
        /// </summary>
        TwoSidedShortEdge,

        /// <summary>
        /// Output prints on both sides of each sheet, which flips along the edge parallel to <see cref="PrintDocumentImageableArea.MediaSizeHeight"/>.
        /// </summary>
        TwoSidedLongEdge
    }

    public enum PageOrder
    {
        /// <summary>
        /// Pages appear in rows, from left to right and top to bottom relative to page orientation.
        /// </summary>
        Horizontal,

        /// <summary>
        /// Pages appear in rows, from right to left and top to bottom relative to page orientation.
        /// </summary>
        HorizontalReverse,

        /// <summary>
        /// Pages appear in columns, from top to bottom and left to right relative to page orientation.
        /// </summary>
        Vertical,

        /// <summary>
        /// Pages appear in columns, from bottom to top and left to right relative to page orientation.
        /// </summary>
        VerticalReverse
    }
}

namespace PrintDialogX.PrintSettings.SettingsHepler
{
    public class NameInfoHepler
    {
        /// <summary>
        /// Initialize a <see cref="NameInfoHepler"/> class.
        /// </summary>
        protected NameInfoHepler()
        {
            return;
        }

        /// <summary>
        /// Get the name info of <see cref="PageMediaSizeName"/> object.
        /// </summary>
        /// <param name="sizeName">The <see cref="PageMediaSizeName"/> object of page size name.</param>
        /// <returns>The name info.</returns>
        public static string GetPageMediaSizeNameInfo(PageMediaSizeName sizeName)
        {
            return sizeName switch
            {
                #region ISO page size: PageMediaSizeName.ISOA4 => "ISO A4",
                PageMediaSizeName.BusinessCard => "Business Card",
                PageMediaSizeName.CreditCard => "Credit Card",
                PageMediaSizeName.ISOA0 => "ISO A0",
                PageMediaSizeName.ISOA1 => "ISO A1",
                PageMediaSizeName.ISOA10 => "ISO A10",
                PageMediaSizeName.ISOA2 => "ISO A2",
                PageMediaSizeName.ISOA3 => "ISO A3",
                PageMediaSizeName.ISOA3Extra => "ISO A3 Extra",
                PageMediaSizeName.ISOA3Rotated => "ISO A3 Rotated",
                PageMediaSizeName.ISOA4 => "ISO A4",
                PageMediaSizeName.ISOA4Extra => "ISO A4 Extra",
                PageMediaSizeName.ISOA4Rotated => "ISO A4 Rotated",
                PageMediaSizeName.ISOA5 => "ISO A5",
                PageMediaSizeName.ISOA5Extra => "ISO A5 Extra",
                PageMediaSizeName.ISOA5Rotated => "ISO A5 Rotated",
                PageMediaSizeName.ISOA6 => "ISO A6",
                PageMediaSizeName.ISOA6Rotated => "ISO A6 Rotated",
                PageMediaSizeName.ISOA7 => "ISO A7",
                PageMediaSizeName.ISOA8 => "ISO A8",
                PageMediaSizeName.ISOA9 => "ISO A9",
                PageMediaSizeName.ISOB0 => "ISO B0",
                PageMediaSizeName.ISOB1 => "ISO B1",
                PageMediaSizeName.ISOB10 => "ISO B10",
                PageMediaSizeName.ISOB2 => "ISO B2",
                PageMediaSizeName.ISOB3 => "ISO B3",
                PageMediaSizeName.ISOB4 => "ISO B4",
                PageMediaSizeName.ISOB4Envelope => "ISO B4 Envelope",
                PageMediaSizeName.ISOB5Envelope => "ISO B5 Envelope",
                PageMediaSizeName.ISOB5Extra => "ISO B5 Extra",
                PageMediaSizeName.ISOB7 => "ISO B7",
                PageMediaSizeName.ISOB8 => "ISO B8",
                PageMediaSizeName.ISOB9 => "ISO B9",
                PageMediaSizeName.ISOC0 => "ISO C0",
                PageMediaSizeName.ISOC1 => "ISO C1",
                PageMediaSizeName.ISOC2 => "ISO C2",
                PageMediaSizeName.ISOC3 => "ISO C3",
                PageMediaSizeName.ISOC3Envelope => "ISO C3 Envelope",
                PageMediaSizeName.ISOC4 => "ISO C4",
                PageMediaSizeName.ISOC4Envelope => "ISO C4 Envelope",
                PageMediaSizeName.ISOC5 => "ISO C5",
                PageMediaSizeName.ISOC5Envelope => "ISO C5 Envelope",
                PageMediaSizeName.ISOC6 => "ISO C6",
                PageMediaSizeName.ISOC6C5Envelope => "ISO C6C5 Envelope",
                PageMediaSizeName.ISOC6Envelope => "ISO C6 Envelope",
                PageMediaSizeName.ISOC7 => "ISO C7",
                PageMediaSizeName.ISOC8 => "ISO C8",
                PageMediaSizeName.ISOC9 => "ISO C9",
                PageMediaSizeName.ISOC10 => "ISO C10",
                PageMediaSizeName.ISODLEnvelope => "ISO DL Envelope",
                PageMediaSizeName.ISODLEnvelopeRotated => "ISO DL Envelope Rotated",
                PageMediaSizeName.ISOSRA3 => "ISO SRA3",
                PageMediaSizeName.Japan2LPhoto => "Japan 2L Photo",
                PageMediaSizeName.JapanChou3Envelope => "Japan Chou 3 Envelope",
                PageMediaSizeName.JapanChou3EnvelopeRotated => "Japan Chou 3 Envelope Rotated",
                PageMediaSizeName.JapanChou4Envelope => "Japan Chou 4 Envelope",
                PageMediaSizeName.JapanChou4EnvelopeRotated => "Japan Chou 4 Envelope Rotated",
                PageMediaSizeName.JapanDoubleHagakiPostcard => "Japan Double Hagaki Postcard",
                PageMediaSizeName.JapanDoubleHagakiPostcardRotated => "Japan Double Hagaki Postcard Rotated",
                PageMediaSizeName.JapanHagakiPostcard => "Japan Hagaki Postcard",
                PageMediaSizeName.JapanHagakiPostcardRotated => "Japan Hagaki Postcard Rotated",
                PageMediaSizeName.JapanKaku2Envelope => "Japan Kaku 2 Envelope",
                PageMediaSizeName.JapanKaku2EnvelopeRotated => "Japan Kaku 2 Envelope Rotated",
                PageMediaSizeName.JapanKaku3Envelope => "Japan Kaku 3 Envelope",
                PageMediaSizeName.JapanKaku3EnvelopeRotated => "Japan Kaku 3 Envelope Rotated",
                PageMediaSizeName.JapanLPhoto => "Japan L Photo",
                PageMediaSizeName.JapanQuadrupleHagakiPostcard => "Japan Quadruple Hagaki Postcard",
                PageMediaSizeName.JapanYou1Envelope => "Japan You 1 Envelope",
                PageMediaSizeName.JapanYou2Envelope => "Japan You 2 Envelope",
                PageMediaSizeName.JapanYou3Envelope => "Japan You 3 Envelope",
                PageMediaSizeName.JapanYou4Envelope => "Japan You 4 Envelope",
                PageMediaSizeName.JapanYou4EnvelopeRotated => "Japan You 4 Envelope Rotated",
                PageMediaSizeName.JapanYou6Envelope => "Japan You 6 Envelope",
                PageMediaSizeName.JapanYou6EnvelopeRotated => "Japan You 6 Envelope Rotated",
                PageMediaSizeName.JISB0 => "JIS B0",
                PageMediaSizeName.JISB1 => "JIS B1",
                PageMediaSizeName.JISB10 => "JIS B10",
                PageMediaSizeName.JISB2 => "JIS B2",
                PageMediaSizeName.JISB3 => "JIS B3",
                PageMediaSizeName.JISB4 => "JIS B4",
                PageMediaSizeName.JISB4Rotated => "JIS B4 Rotated",
                PageMediaSizeName.JISB5 => "JIS B5",
                PageMediaSizeName.JISB5Rotated => "JIS B5 Rotated",
                PageMediaSizeName.JISB6 => "JIS B6",
                PageMediaSizeName.JISB6Rotated => "JIS B6 Rotated",
                PageMediaSizeName.JISB7 => "JIS B7",
                PageMediaSizeName.JISB8 => "JIS B8",
                PageMediaSizeName.JISB9 => "JIS B9",
                PageMediaSizeName.NorthAmerica10x11 => "North America 10 x 11",
                PageMediaSizeName.NorthAmerica10x12 => "North America 10 x 12",
                PageMediaSizeName.NorthAmerica10x14 => "North America 10 x 14",
                PageMediaSizeName.NorthAmerica11x17 => "North America 11 x 17",
                PageMediaSizeName.NorthAmerica14x17 => "North America 14 x 17",
                PageMediaSizeName.NorthAmerica4x6 => "North America 4 x 6",
                PageMediaSizeName.NorthAmerica4x8 => "North America 4 x 8",
                PageMediaSizeName.NorthAmerica5x7 => "North America 5 x 7",
                PageMediaSizeName.NorthAmerica8x10 => "North America 8 x 10",
                PageMediaSizeName.NorthAmerica9x11 => "North America 9 x 11",
                PageMediaSizeName.NorthAmericaArchitectureASheet => "North America Architecture A Sheet",
                PageMediaSizeName.NorthAmericaArchitectureBSheet => "North America Architecture B Sheet",
                PageMediaSizeName.NorthAmericaArchitectureCSheet => "North America Architecture C Sheet",
                PageMediaSizeName.NorthAmericaArchitectureDSheet => "North America Architecture D Sheet",
                PageMediaSizeName.NorthAmericaArchitectureESheet => "North America Architecture E Sheet",
                PageMediaSizeName.NorthAmericaCSheet => "North America C Sheet",
                PageMediaSizeName.NorthAmericaDSheet => "North America D Sheet",
                PageMediaSizeName.NorthAmericaESheet => "North America E Sheet",
                PageMediaSizeName.NorthAmericaExecutive => "North America Executive",
                PageMediaSizeName.NorthAmericaGermanLegalFanfold => "North America German Legal Fanfold",
                PageMediaSizeName.NorthAmericaGermanStandardFanfold => "North America German Standard Fanfold",
                PageMediaSizeName.NorthAmericaLegal => "North America Legal",
                PageMediaSizeName.NorthAmericaLegalExtra => "North America Legal Extra",
                PageMediaSizeName.NorthAmericaLetter => "North America Letter",
                PageMediaSizeName.NorthAmericaLetterExtra => "North America Letter Extra",
                PageMediaSizeName.NorthAmericaLetterPlus => "North America Letter Plus",
                PageMediaSizeName.NorthAmericaLetterRotated => "North America Letter Rotated",
                PageMediaSizeName.NorthAmericaMonarchEnvelope => "North America Monarch Envelope",
                PageMediaSizeName.NorthAmericaNote => "North America Note",
                PageMediaSizeName.NorthAmericaNumber10Envelope => "North America Number 10 Envelope",
                PageMediaSizeName.NorthAmericaNumber10EnvelopeRotated => "North America Number 10 Envelope Rotated",
                PageMediaSizeName.NorthAmericaNumber11Envelope => "North America Number 11 Envelope",
                PageMediaSizeName.NorthAmericaNumber12Envelope => "North America Number 12 Envelope",
                PageMediaSizeName.NorthAmericaNumber14Envelope => "North America Number 14 Envelope",
                PageMediaSizeName.NorthAmericaNumber9Envelope => "North America Number 9 Envelope",
                PageMediaSizeName.NorthAmericaPersonalEnvelope => "North America Personal Envelope",
                PageMediaSizeName.NorthAmericaQuarto => "North America Quarto",
                PageMediaSizeName.NorthAmericaStatement => "North America Statement",
                PageMediaSizeName.NorthAmericaSuperA => "North America Super A",
                PageMediaSizeName.NorthAmericaSuperB => "North America Super B",
                PageMediaSizeName.NorthAmericaTabloid => "North America Tabloid",
                PageMediaSizeName.NorthAmericaTabloidExtra => "North America Tabloid Extra",
                PageMediaSizeName.OtherMetricA3Plus => "A3 Plus",
                PageMediaSizeName.OtherMetricA4Plus => "A4 Plus",
                PageMediaSizeName.OtherMetricFolio => "Folio",
                PageMediaSizeName.OtherMetricInviteEnvelope => "Invite Envelope",
                PageMediaSizeName.OtherMetricItalianEnvelope => "Italian Envelope",
                PageMediaSizeName.PRC10Envelope => "PRC #10 Envelope",
                PageMediaSizeName.PRC10EnvelopeRotated => "PRC #10 Envelope Rotated",
                PageMediaSizeName.PRC16K => "PRC 16K",
                PageMediaSizeName.PRC16KRotated => "PRC 16K Rotated",
                PageMediaSizeName.PRC1Envelope => "PRC #1 Envelope",
                PageMediaSizeName.PRC1EnvelopeRotated => "PRC #1 Envelope Rotated",
                PageMediaSizeName.PRC2Envelope => "PRC #2 Envelope",
                PageMediaSizeName.PRC2EnvelopeRotated => "PRC #2 Envelope Rotated",
                PageMediaSizeName.PRC32K => "PRC 32K",
                PageMediaSizeName.PRC32KBig => "PRC 32K Big",
                PageMediaSizeName.PRC32KRotated => "PRC 32K Rotated",
                PageMediaSizeName.PRC3Envelope => "PRC #3 Envelope",
                PageMediaSizeName.PRC3EnvelopeRotated => "PRC #3 Envelope Rotated",
                PageMediaSizeName.PRC4Envelope => "PRC #4 Envelope",
                PageMediaSizeName.PRC4EnvelopeRotated => "PRC #4 Envelope Rotated",
                PageMediaSizeName.PRC5Envelope => "PRC #5 Envelope",
                PageMediaSizeName.PRC5EnvelopeRotated => "PRC #5 Envelope Rotated",
                PageMediaSizeName.PRC6Envelope => "PRC #6 Envelope",
                PageMediaSizeName.PRC6EnvelopeRotated => "PRC #6 Envelope Rotated",
                PageMediaSizeName.PRC7Envelope => "PRC #7 Envelope",
                PageMediaSizeName.PRC7EnvelopeRotated => "PRC #7 Envelope Rotated",
                PageMediaSizeName.PRC8Envelope => "PRC #8 Envelope",
                PageMediaSizeName.PRC8EnvelopeRotated => "PRC #8 Envelope Rotated",
                PageMediaSizeName.PRC9Envelope => "PRC #9 Envelope",
                PageMediaSizeName.PRC9EnvelopeRotated => "PRC #9 Envelope Rotated",
                PageMediaSizeName.Roll04Inch => "4-inch Wide Roll",
                PageMediaSizeName.Roll06Inch => "6-inch Wide Roll",
                PageMediaSizeName.Roll08Inch => "8-inch Wide Roll",
                PageMediaSizeName.Roll12Inch => "12-inch Wide Roll",
                PageMediaSizeName.Roll15Inch => "15-inch Wide Roll",
                PageMediaSizeName.Roll18Inch => "18-inch Wide Roll",
                PageMediaSizeName.Roll22Inch => "22-inch Wide Roll",
                PageMediaSizeName.Roll24Inch => "24-inch Wide Roll",
                PageMediaSizeName.Roll30Inch => "30-inch Wide Roll",
                PageMediaSizeName.Roll36Inch => "36-inch Wide Roll",
                PageMediaSizeName.Roll54Inch => "54-inch Wide Roll",

                #endregion

                _ => "Unknown Size",
            };
        }

        /// <summary>
        /// Get the name info of <see cref="PrintSettings.PageSize"/> object.
        /// </summary>
        /// <param name="sizeName">The <see cref="PrintSettings.PageSize"/> object of page size name.</param>
        /// <returns>The name info.</returns>

        /// <summary>
        /// Get the name info of <see cref="PageMediaType"/> object.
        /// </summary>
        /// <param name="type">The <see cref="PageMediaType"/> object of page type.</param>
        /// <returns>The name info.</returns>
        public static string GetPageMediaTypeNameInfo(PageMediaType type)
        {
            return type switch
            {
                PageMediaType.Archival => "Archival",
                PageMediaType.AutoSelect => "Auto Select",
                PageMediaType.BackPrintFilm => "Back Print Film",
                PageMediaType.Bond => "Bond",
                PageMediaType.CardStock => "Card Stock",
                PageMediaType.Continuous => "Continuous",
                PageMediaType.EnvelopePlain => "Envelope Plain",
                PageMediaType.EnvelopeWindow => "Envelope Window",
                PageMediaType.Fabric => "Fabric",
                PageMediaType.HighResolution => "High Resolution",
                PageMediaType.Label => "Label",
                PageMediaType.MultiLayerForm => "Multi Layer Form",
                PageMediaType.MultiPartForm => "Multi Part Form",
                PageMediaType.Photographic => "Photographic",
                PageMediaType.PhotographicFilm => "Photographic Film",
                PageMediaType.PhotographicGlossy => "Photographic Glossy",
                PageMediaType.PhotographicHighGloss => "Photographic High Gloss",
                PageMediaType.PhotographicMatte => "Photographic Matte",
                PageMediaType.PhotographicSatin => "Photographic Satin",
                PageMediaType.PhotographicSemiGloss => "Photographic Semi Gloss",
                PageMediaType.Plain => "Plain",
                PageMediaType.Screen => "Screen",
                PageMediaType.ScreenPaged => "Screen Paged",
                PageMediaType.Stationery => "Stationery",
                PageMediaType.TabStockFull => "Tab Stock Full",
                PageMediaType.TabStockPreCut => "Tab Stock Pre Cut",
                PageMediaType.Transparency => "Transparency",
                PageMediaType.TShirtTransfer => "T-shirt Transfer",

                _ => "Unknown Type",
            };
        }

        /// <summary>
        /// Get the name info of <see cref="PrintSettings.PageType"/> object.
        /// </summary>
        /// <param name="type">The <see cref="PrintSettings.PageType"/> object of page type.</param>
        /// <returns>The name info.</returns>
        public static string GetPageMediaTypeNameInfo(PrintSettings.PageType type)
        {
            return type switch
            {
                PrintSettings.PageType.Archival => "Archival",
                PrintSettings.PageType.AutoSelect => "Auto Select",
                PrintSettings.PageType.BackPrintFilm => "Back Print Film",
                PrintSettings.PageType.Bond => "Bond",
                PrintSettings.PageType.CardStock => "Card Stock",
                PrintSettings.PageType.Continuous => "Continuous",
                PrintSettings.PageType.EnvelopePlain => "Envelope Plain",
                PrintSettings.PageType.EnvelopeWindow => "Envelope Window",
                PrintSettings.PageType.Fabric => "Fabric",
                PrintSettings.PageType.HighResolution => "High Resolution",
                PrintSettings.PageType.Label => "Label",
                PrintSettings.PageType.MultiLayerForm => "Multi Layer Form",
                PrintSettings.PageType.MultiPartForm => "Multi Part Form",
                PrintSettings.PageType.Photographic => "Photographic",
                PrintSettings.PageType.PhotographicFilm => "Photographic Film",
                PrintSettings.PageType.PhotographicGlossy => "Photographic Glossy",
                PrintSettings.PageType.PhotographicHighGloss => "Photographic High Gloss",
                PrintSettings.PageType.PhotographicMatte => "Photographic Matte",
                PrintSettings.PageType.PhotographicSatin => "Photographic Satin",
                PrintSettings.PageType.PhotographicSemiGloss => "Photographic Semi Gloss",
                PrintSettings.PageType.Plain => "Plain",
                PrintSettings.PageType.Screen => "Screen",
                PrintSettings.PageType.ScreenPaged => "Screen Paged",
                PrintSettings.PageType.Stationery => "Stationery",
                PrintSettings.PageType.TabStockFull => "Tab Stock Full",
                PrintSettings.PageType.TabStockPreCut => "Tab Stock Pre Cut",
                PrintSettings.PageType.Transparency => "Transparency",
                PrintSettings.PageType.TShirtTransfer => "T-shirt Transfer",

                _ => "Unknown Type",
            };
        }

        /// <summary>
        /// Get the name info of <see cref="InputBin"/> object.
        /// </summary>
        /// <param name="inputBin">The <see cref="InputBin"/> object of page source.</param>
        /// <returns>The name info.</returns>
        public static string GetInputBinNameInfo(InputBin inputBin)
        {
            return inputBin switch
            {
                InputBin.AutoSelect => "Auto Select",
                InputBin.AutoSheetFeeder => "Auto Sheet Feeder",
                InputBin.Cassette => "Cassette",
                InputBin.Manual => "Manual",
                InputBin.Tractor => "Tractor",

                _ => "Unknown Input Bin",
            };
        }
    }
}


