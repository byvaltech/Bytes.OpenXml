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
    public static WorksheetPart GetWorksheetPart(this WorkbookPart workbookPart, string sheetName)
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
        return (WorksheetPart)workbookPart.GetPartById(relationshipId);
    }

    /// <summary>
    /// Get cell format for the given id
    /// </summary>
    /// <param name="workbookPart">An instance of <see cref="WorkbookPart"/></param>
    /// <param name="numberFormatId">Number format ID</param>
    /// <exception cref="ArgumentNullException"></exception>
    public static uint GetCellFormat(this WorkbookPart workbookPart, uint numberFormatId)
    {
        if (workbookPart == null)
        {
            throw new ArgumentNullException(nameof(workbookPart));
        }

        var styleSheet = workbookPart.WorkbookStylesPart.Stylesheet;
        var elements = styleSheet.CellFormats.Elements<CellFormat>().ToList();
        var foundItem = elements.Find(e => e.NumberFormatId == numberFormatId);
        if (foundItem != null)
        {
            return (uint)elements.IndexOf(foundItem);
        }
        return 0;
    }
}