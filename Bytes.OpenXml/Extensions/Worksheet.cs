namespace Bytes.OpenXml.Extensions;

/// <summary>
/// Extensions for <see cref="Worksheet"/>
/// </summary>
public static class WorksheetExtensions
{
    /// <summary>
    /// Get a row by index
    /// </summary>
    /// <param name="worksheet">An instance of <see cref="Worksheet"/>.</param>
    /// <param name="rowIndex">Index of the required row</param>
    public static Row GetRow(this Worksheet worksheet, uint rowIndex)
    {
        if (worksheet == null)
        {
            throw new ArgumentNullException(nameof(worksheet));
        }

        return worksheet.GetFirstChild<SheetData>().Elements<Row>().First(r => r.RowIndex == rowIndex);
    }

    /// <summary>
    /// Get a cell by column name and row index
    /// </summary>
    /// <param name="worksheet">An instance of <see cref="Worksheet"/>.</param>
    /// <param name="columnName">Name of the column</param>
    /// <param name="rowIndex">Index of the required row</param>
    public static Cell GetCell(this Worksheet worksheet, string columnName, uint rowIndex)
    {
        if (worksheet == null)
        {
            throw new ArgumentNullException(nameof(worksheet));
        }

        var row = worksheet.GetRow(rowIndex);

        if (row == null)
        {
            return null;
        }

        return row.Elements<Cell>().First(c => string.Equals(c.CellReference.Value, columnName + rowIndex, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Find cells containing given text
    /// </summary>
    /// <param name="worksheet">An instance of <see cref="Worksheet"/></param>
    /// <param name="search">Text to search</param>
    public static IList<(uint row, uint column)> FindCellsContaining(this Worksheet worksheet, string search)
    {
        var indexes = new List<(uint, uint)>();

        if (worksheet != null)
        {
            foreach (var row in worksheet.Descendants<Row>().ToList())
            {
                foreach (var cell in row.Elements<Cell>().ToList())
                {
                    if (cell.InnerText.IndexOf(search, StringComparison.OrdinalIgnoreCase) > -1)
                    {
                        indexes.Add((row.RowIndex.Value, cell.GetColumnIndex()));
                    }
                }
            }
        }
        return indexes;
    }

    /// <summary>
    /// Resize a table
    /// </summary>
    /// <param name="worksheet">An instance of <see cref="Worksheet"/></param>
    /// <param name="tableName">Name of the table</param>
    /// <param name="rows">No. of rows</param>
    /// <exception cref="ArgumentNullException"></exception>
    /// <exception cref="ArgumentOutOfRangeException"></exception>
    public static void ResizeTable(this Worksheet worksheet, string tableName, uint rows)
    {
        if (worksheet == null)
        {
            throw new ArgumentNullException(nameof(worksheet));
        }

        worksheet.WorksheetPart.ResizeTable(tableName, rows);
    }

    /// <summary>
    /// Get a table by name
    /// </summary>
    /// <param name="worksheet"></param>
    /// <param name="tableName"></param>
    public static TableDefinitionPart GetTable(this Worksheet worksheet, string tableName)
    {
        if (worksheet == null)
        {
            throw new ArgumentNullException(nameof(worksheet));
        }

        return worksheet.WorksheetPart.GetTable(tableName);
    }

    /// <summary>
    /// Refresh Pivots on load
    /// </summary>
    /// <param name="worksheet">An instance of <see cref="Worksheet"/></param>
    /// <exception cref="ArgumentNullException"></exception>
    /// <exception cref="ArgumentOutOfRangeException"></exception>
    public static void RefreshPivotsOnLoad(this Worksheet worksheet)
    {
        if (worksheet == null)
        {
            throw new ArgumentNullException(nameof(worksheet));
        }
        worksheet.WorksheetPart.RefreshPivotsOnLoad();
    }

    /// <summary>
    /// Refresh Pivot on load
    /// </summary>
    /// <param name="worksheet">An instance of <see cref="Worksheet"/></param>
    /// <param name="name">Name of the pivot to refresh</param>
    /// <exception cref="ArgumentNullException"></exception>
    /// <exception cref="ArgumentOutOfRangeException"></exception>
    public static void RefreshPivotOnLoad(this Worksheet worksheet, string name)
    {
        if (worksheet == null)
        {
            throw new ArgumentNullException(nameof(worksheet));
        }

        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException("Cannot be null or empty", nameof(name));
        }

        worksheet.WorksheetPart.RefreshPivotOnLoad(name);
    }
}