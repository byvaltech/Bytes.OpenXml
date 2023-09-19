namespace Bytes.OpenXml.Extensions;

/// <summary>
/// Extensions for <see cref="WorkbookPart"/>
/// </summary>
public static class WorkbookPartExtensions
{
    /// <summary>
    /// Get worksheet part by sheet name
    /// </summary>
    /// <param name="workbookPart">An instance of <see cref="WorkbookPart"/>.</param>
    /// <param name="sheetName">Name of the sheet</param>
    /// <returns></returns>
    public static WorksheetPart GetWorksheetPartByName(this WorkbookPart workbookPart, string sheetName)
    {
        if (workbookPart == null)
        {
            throw new ArgumentNullException(nameof(workbookPart));
        }

        var sheets = workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);

        if (!sheets.Any())
        {
            return null;
        }

        var relationshipId = sheets.First().Id.Value;
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(relationshipId);

        return worksheetPart;
    }
}