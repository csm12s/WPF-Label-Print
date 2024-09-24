
using System.ComponentModel.DataAnnotations;

namespace PrintDialogX.Test;

public class PartLabelDto : LabelDtoBase
{
    public string? PartRef { get; set; }
    public string? PartName { get; set; }

    public string? MatRef { get; set; }


    public double? Thickness { get; set; }

    public string? Remark { get; set; }
}

