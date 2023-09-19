namespace Bytes.OpenXml.Extensions;

/// <summary>
/// Extensions for <see cref="SpreadsheetDocument"/>
/// </summary>
public static class SpreadsheetDocumentExtensions
{
    /// <summary>
    /// Get worksheet part by sheet name
    /// </summary>
    /// <param name="document">An instance of <see cref="SpreadsheetDocument"/>.</param>
    /// <param name="sheetName">Name of the sheet</param>
    /// <returns></returns>
    public static WorksheetPart GetWorksheetPartByName(this SpreadsheetDocument document, string sheetName)
    {
        if (document == null)
        {
            throw new ArgumentNullException(nameof(document));
        }
        return document.WorkbookPart.GetWorksheetPartByName(sheetName);
    }
}