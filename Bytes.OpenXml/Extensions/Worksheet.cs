using System.Text.RegularExpressions;

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

        var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>();

        if (rows.Any())
        {
            return worksheet.GetFirstChild<SheetData>().Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
        }
        return null;
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
    /// <param name="workbookPart"></param>
    /// <param name="search">Text to search</param>
    public static IList<(uint row, uint column)> FindCellsContaining(this Worksheet worksheet, WorkbookPart workbookPart, string search)
    {
        var indexes = new List<(uint, uint)>();

        if (worksheet != null && workbookPart != null)
        {
            foreach (var row in worksheet.Descendants<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    var value = cell.InnerText;
                    var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    var isInt = int.TryParse(value, out var valueno);
                    if (isInt)
                    {
                        value = stringTable.SharedStringTable.ElementAt(valueno).InnerText;
                    }
                    if (value.IndexOf(search, StringComparison.OrdinalIgnoreCase) > -1)
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

    /// <summary>
    /// Remove a row at given position
    /// </summary>
    /// <param name="worksheet">An instance of <see cref="Worksheet"/></param>
    /// <param name="rowIndex">Index of the row to remove</param>
    /// <exception cref="ArgumentNullException"></exception>
    public static void RemoveRowAt(this Worksheet worksheet, int rowIndex)
    {
        if (worksheet == null)
        {
            throw new ArgumentNullException(nameof(worksheet));
        }
        var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>();

        var matched = false;
        foreach (var row in rows)
        {
            if (matched)
            {
                row.RowIndex.Value--;

                IEnumerable<Cell> cells = row.Elements<Cell>().ToList();
                if (cells != null)
                {
                    foreach (Cell cell in cells)
                    {
                        var cr = cell.CellReference.Value;

                        int indexRow = Convert.ToInt32(Regex.Replace(cr, @"[^\d]+", "")) - 1;
                        cr = Regex.Replace(cr, @"[\d-]", "");

                        cell.CellReference.Value = $"{cr}{indexRow}";
                    }
                }
            }
            if (row.RowIndex == rowIndex)
            {
                matched = true;
                row.Remove();
            }
        }
    }
}