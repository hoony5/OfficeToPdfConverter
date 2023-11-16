using Microsoft.Office.Interop.Excel;

public class ExcelToPdfConverter
{
    public void Convert(string excelPath, string saveAs)
    {
        Application excel = new Application();
        Workbook wb = excel.Workbooks.Open(excelPath);
        wb.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, saveAs);
        wb.Close(false);
        excel.Quit();
        // check platform only window
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            _ = Marshal.FinalReleaseComObject(wb);
    }
}