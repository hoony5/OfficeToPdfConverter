using Microsoft.Office.Interop.PowerPoint;

public class PowerPointToPdfConverter
{
    public void Convert(string powerPointPath, string saveAs)
    {
        Application powerPoint = new Application();
        Presentation ppt = powerPoint.Presentations.Open(powerPointPath);
        ppt.ExportAsFixedFormat(saveAs, PpFixedFormatType.ppFixedFormatTypePDF);
        ppt.Close();
        powerPoint.Quit();
        // check platform only window
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            _ = Marshal.FinalReleaseComObject(ppt);
    }
}