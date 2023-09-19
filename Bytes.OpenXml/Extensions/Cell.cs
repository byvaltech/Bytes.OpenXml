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
}