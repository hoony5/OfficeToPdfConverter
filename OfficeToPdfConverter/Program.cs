// See https://aka.ms/new-console-template for more information


while(true)
{
    string? command = Console.ReadLine();
    if (string.IsNullOrEmpty(command)) continue;
    if(command.Contains("Exit", StringComparison.OrdinalIgnoreCase)) break;
    CommandProcessor commandProcessor = new();
    await commandProcessor.ProcessCommand(command);
}

