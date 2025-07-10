using ExcelReader.Model;
using Spectre.Console;

namespace ExcelReader.View;

public class Display
{
    public static void PrintAllData(string tableName)
    {
        var repository = new DataRepository(DatabaseUtility.GetConnectionString());
        var data = repository.GetAllData(tableName);

        var rule = new Rule($"[green]{tableName}[/]");
        rule.Justification = Justify.Left;
        AnsiConsole.Write(rule);

        if (data == null || data.Count == 0)
        {
            AnsiConsole.MarkupLine("[red]No records found.[/]");
            return;
        }

        var table = new Table();

        foreach (var key in data[0].Keys)
        {
            table.AddColumn(new TableColumn($"[dodgerblue1]{key}[/]").Centered().NoWrap());
        }

        foreach (var row in data)
        {
            var rowValues = row.Values.Select(value => value?.ToString() ?? string.Empty).ToArray();
            table.AddRow(rowValues);
        }

        int maxWidth = Console.LargestWindowWidth;
        int maxHeight = Console.LargestWindowHeight;
        int desiredWidth = Math.Min(200, maxWidth);
        int desiredHeight = Math.Min(50, maxHeight);
        Console.SetWindowSize(desiredWidth, desiredHeight);

        AnsiConsole.Write(table.Width(190).ShowRowSeparators().Border(TableBorder.DoubleEdge));
        System.Threading.Thread.Sleep(2000);
    }
}
