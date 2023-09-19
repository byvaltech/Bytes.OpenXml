namespace Bytes.OpenXml.Extensions;

/// <summary>
/// Extensions for <see cref="WorksheetPart"/>
/// </summary>
public static class WorksheetPartExtensions
{
    /// <summary>
    /// Resize a table
    /// </summary>
    /// <param name="worksheetPart">An instance of <see cref="WorksheetPart"/></param>
    /// <param name="tableName">Name of the table</param>
    /// <param name="rows">No. of rows</param>
    /// <exception cref="ArgumentNullException"></exception>
    /// <exception cref="ArgumentOutOfRangeException"></exception>
    public static void ResizeTable(this WorksheetPart worksheetPart, string tableName, uint rows)
    {
        if (worksheetPart == null)
        {
            throw new ArgumentNullException(nameof(worksheetPart));
        }

        var tableDefinitionPart = worksheetPart.GetTable(tableName) ?? throw new ArgumentOutOfRangeException(nameof(tableName), "Table not found");

        var newSize = tableDefinitionPart.Table.UpdateRowsTo(rows);
        tableDefinitionPart.Table.Reference = newSize;
        tableDefinitionPart.Table.AutoFilter.Reference = newSize;
    }

    /// <summary>
    /// Get a table by name
    /// </summary>
    /// <param name="worksheetPart"></param>
    /// <param name="tableName"></param>
    public static TableDefinitionPart GetTable(this WorksheetPart worksheetPart, string tableName)
    {
        if (worksheetPart == null)
        {
            throw new ArgumentNullException(nameof(worksheetPart));
        }

        return worksheetPart.TableDefinitionParts.FirstOrDefault(table => string.Equals(table.Table.DisplayName, tableName, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Refresh Pivots on load
    /// </summary>
    /// <param name="worksheetPart">An instance of <see cref="WorksheetPart"/></param>
    /// <exception cref="ArgumentNullException"></exception>
    /// <exception cref="ArgumentOutOfRangeException"></exception>
    public static void RefreshPivotsOnLoad(this WorksheetPart worksheetPart)
    {
        if (worksheetPart == null)
        {
            throw new ArgumentNullException(nameof(worksheetPart));
        }

        foreach (var pivot in worksheetPart.PivotTableParts)
        {
            pivot.PivotTableCacheDefinitionPart.PivotCacheDefinition.RefreshOnLoad = true;
        }
    }

    /// <summary>
    /// Refresh Pivot on load
    /// </summary>
    /// <param name="worksheetPart">An instance of <see cref="WorksheetPart"/></param>
    /// <param name="name">Name of the pivot to refresh</param>
    /// <exception cref="ArgumentNullException"></exception>
    /// <exception cref="ArgumentOutOfRangeException"></exception>
    public static void RefreshPivotOnLoad(this WorksheetPart worksheetPart, string name)
    {
        if (worksheetPart == null)
        {
            throw new ArgumentNullException(nameof(worksheetPart));
        }

        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException("Cannot be null or empty", nameof(name));
        }

        var pivot = worksheetPart.PivotTableParts.FirstOrDefault(p => string.Equals(p.PivotTableDefinition.Name, name, StringComparison.OrdinalIgnoreCase));
        if (pivot != null)
        {
            pivot.PivotTableCacheDefinitionPart.PivotCacheDefinition.RefreshOnLoad = true;
        }
    }
}