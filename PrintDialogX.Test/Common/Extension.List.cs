using Newtonsoft.Json;
using System.Collections;
using System.Data;

namespace PrintDialogX.Common;

/// <summary>
/// List / IEnumerable extension
/// </summary>
public static partial class Extension
{
    /// <summary>
    /// class to data table
    /// </summary>
    /// <param name="list"></param>
    /// <returns></returns>
    public static DataTable ToDataTable(this IEnumerable list)
    {
        var json = JsonConvert.SerializeObject(list);
        DataTable table = (DataTable)JsonConvert.DeserializeObject(json, typeof(DataTable));
        return table;

        // For performance see: https://stackoverflow.com/questions/564366/convert-generic-list-enumerable-to-datatable
    }


    /// <summary>
    /// Has Items
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="list"></param>
    /// <returns></returns>
    public static bool HasValue<T>(this IEnumerable<T> list)
    {
        var hasValue = false;

        if (list != null && list.Any())
        {
            hasValue = true;
        }

        return hasValue;
    }
}
