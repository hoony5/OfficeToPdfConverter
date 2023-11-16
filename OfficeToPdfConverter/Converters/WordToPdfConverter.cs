using Microsoft.Office.Interop.Word;

public class WordToPdfConverter
{
    public void Convert(string wordPath, string saveAs)
    {
        Application word = new Application();
        Document doc = word.Documents.Open(wordPath);
        doc.ExportAsFixedFormat(saveAs, WdExportFormat.wdExportFormatPDF);
        doc.Close(false);
        word.Quit();
        // check platform only window
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            _ = Marshal.FinalReleaseComObject(doc);
    }
}