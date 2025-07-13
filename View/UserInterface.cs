namespace ExcelReader.View;

public class UserInterface
{
    public string GetFilePath()
    {
        Console.WriteLine("Please enter the full path of the Excel, CSV, or PDF file (e.g., C:\\path\\to\\your\\file.xlsx [or .csv or .pdf]):");
        string filePath = Console.ReadLine();

        Console.WriteLine("\nLooking for the file at: " + filePath + "...");

        while (!File.Exists(filePath))
        {
            Console.WriteLine("File not found. \nPlease enter a valid file path:");
            filePath = Console.ReadLine();
            Console.WriteLine("\nLooking for the file at: " + filePath + "...");
        }

        Console.WriteLine("File found.\n");

        return filePath;
    }
}
