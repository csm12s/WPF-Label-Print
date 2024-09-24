namespace PrintDialogX.Test;


// each label has 1 of below items, in xaml set the name as: <Image x:Name="Image" 
public class LabelDtoBase
{

    // image
    public string? Image { get; set; }

    // todo not show 
    // bar cdoe
    public string? Barcode { get; set; }

    // qr code
    public string? QRCode { get; set; }

    // if there is 100 parts, set the "比例数量" (percent) like 10%, it will print 10 labels for the part, by default print 1 label
    public int? Quantity { get; set; }

}