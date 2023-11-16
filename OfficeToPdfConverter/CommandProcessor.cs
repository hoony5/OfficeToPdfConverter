public class CommandProcessor
{
    private const char SPLIT_SIGNATURE = '-';
    private const char SUB_SPLIT_SIGNATURE = '=';
    private const string EXCEL_KEYWORD = "Excel";
    private const string WORD_KEYWORD = "Word";
    private const string POWER_POINT_KEYWORD = "PowerPoint";
    private const string FILE_TYPE_SIGNATURE = "Type";
    private const string FILE_PATH_SIGNATURE = "Path";
    private const string SAVE_PATH_SIGNATURE = "Save";
    private const int KEYWORD_SIGNATURE_INDEX = 1;
    
    public async Task ProcessCommand(string commandLine)
    {
        string[] split = commandLine.Split(SPLIT_SIGNATURE);
        // subSplit Pattern => -type=Word -path=... -save=...
        // subSplit Pattern => -type=Excel -path=... -save=...
        // subSplit Pattern => -type=PowerPoint -path=... -save=...
        string? filePath = Array.Find(split,
            line => line
                    .Contains(FILE_PATH_SIGNATURE, StringComparison.OrdinalIgnoreCase))?
                    .Split(SUB_SPLIT_SIGNATURE)[KEYWORD_SIGNATURE_INDEX]
                    .Trim();
        
        string? saveAs = Array.Find(split,
            line => line
                    .Contains(SAVE_PATH_SIGNATURE, StringComparison.OrdinalIgnoreCase))?
                    .Split(SUB_SPLIT_SIGNATURE)[KEYWORD_SIGNATURE_INDEX]
                    .Trim();
        
        string? fileType = Array.Find(split,
            line => line
                    .Contains(FILE_TYPE_SIGNATURE, StringComparison.OrdinalIgnoreCase))?
                    .Split(SUB_SPLIT_SIGNATURE)[KEYWORD_SIGNATURE_INDEX]
                    .Trim();
        
        if (fileType is null || saveAs is null || filePath is null) return;
        
        await Task.Run(() => 
        {
            if (IsExcelFile(fileType)) 
                new ExcelToPdfConverter().Convert(filePath, saveAs);
            else if (IsWordFile(fileType)) 
                new WordToPdfConverter().Convert(filePath, saveAs);
            else if (IsPowerPointFile(fileType))
                new PowerPointToPdfConverter().Convert(filePath, saveAs);
        }, CancellationToken.None);
    }

    private bool IsExcelFile(string input)
    {
        return input.Contains(EXCEL_KEYWORD, StringComparison.OrdinalIgnoreCase);
    }
    private bool IsWordFile(string input)
    {
        return input.Contains(WORD_KEYWORD, StringComparison.OrdinalIgnoreCase);
    }
    private bool IsPowerPointFile(string input)
    {
        return input.Contains(POWER_POINT_KEYWORD, StringComparison.OrdinalIgnoreCase);
    }
}
