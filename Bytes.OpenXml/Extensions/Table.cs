using System.Text.RegularExpressions;

namespace Bytes.OpenXml.Extensions;

/// <summary>
/// Extensions for <see cref="Table"/>
/// </summary>
public static class TableExtensions
{
    /// <summary>
    /// Updates the rows of table reference to given number of rows.
    /// </summary>
    /// <param name="table">An instance of <see cref="Table"/></param>
    /// <param name="rows">No of rows</param>
    /// <returns>size after updating the table</returns>
    public static string UpdateRowsTo(this Table table, uint rows)
    {
        if (table == null)
        {
            throw new ArgumentNullException(nameof(table));
        }
        var result = table.Reference.Value.Trim();
        var parts = result.Split(':');
        var regex = new Regex("[a-zA-Z]+");
        var match = regex.Match(parts[1]);
        return $"{parts[0]}:{match.Value}{rows}";
    }
}