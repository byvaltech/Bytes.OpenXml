using System.Globalization;
using System.Text.RegularExpressions;

namespace Bytes.OpenXml.Extensions;

/// <summary>
/// Extensions for <see cref="Cell"/>
/// </summary>
public static class CellExtensions
{
    /// <summary>
    /// Get column index of a <see cref="Cell"/>
    /// </summary>
    /// <param name="cell">An instance of <see cref="Cell"/></param>
    public static uint GetColumnIndex(this Cell cell)
    {
        if (cell == null)
        {
            throw new ArgumentNullException(nameof(cell));
        }

        var columnReference = Regex.Replace(cell.CellReference.ToString().ToUpperInvariant(), @"[\d]", string.Empty);

        var columnNumber = -1;
        var mulitplier = 1;

        foreach (var c in columnReference.ToCharArray().Reverse())
        {
            columnNumber += mulitplier * (c - 64);

            mulitplier *= 26;
        }

        return (uint)(columnNumber + 1);
    }

    /// <summary>
    /// Get Row Index of the cell
    /// </summary>
    /// <param name="cell"></param>
    public static uint GetRowIndex(this Cell cell)
    {
        if (cell == null)
        {
            throw new ArgumentNullException(nameof(cell));
        }
        var regex = new Regex(@"\d+");
        var match = regex.Match(cell.CellReference);

        return uint.Parse(match.Value, CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Get Column name of the cell
    /// </summary>
    /// <param name="cell"></param>
    public static string GetColumnName(this Cell cell)
    {
        if (cell == null)
        {
            throw new ArgumentNullException(nameof(cell));
        }

        var regex = new Regex("[A-Za-z]+");
        var match = regex.Match(cell.CellReference);

        return match.Value;
    }
}