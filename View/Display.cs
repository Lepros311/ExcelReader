using ExcelReader.Model;
using Spectre.Console;

namespace ExcelReader.View;

public class Display
{
    public static void PrintAllContacts(string heading)
    {
        var repository = new ContactsRepository(DatabaseUtility.GetConnectionString());
        var contacts = repository.GetAllContacts();

        var rule = new Rule($"[green]{heading}[/]");
        rule.Justification = Justify.Left;
        AnsiConsole.Write(rule);

        var table = new Table()
            .AddColumn(new TableColumn("[dodgerblue1]ID[/]").Centered().Width(5).NoWrap())
            .AddColumn(new TableColumn("[dodgerblue1]First Name[/]").Centered().Width(15).NoWrap())
            .AddColumn(new TableColumn("[dodgerblue1]Last Name[/]").Centered().Width(20).NoWrap())
            .AddColumn(new TableColumn("[dodgerblue1]Phone Number[/]").Centered().Width(20).NoWrap())
            .AddColumn(new TableColumn("[dodgerblue1]Email Address[/]").Centered().Width(30).NoWrap())
            .AddColumn(new TableColumn("[dodgerblue1]Address Line 1[/]").Centered().Width(25).NoWrap())
            .AddColumn(new TableColumn("[dodgerblue1]Address Line 2[/]").Centered().Width(20).NoWrap())
            .AddColumn(new TableColumn("[dodgerblue1]City[/]").Centered().Width(20).NoWrap())
            .AddColumn(new TableColumn("[dodgerblue1]State[/]").Centered().Width(20).NoWrap())
            .AddColumn(new TableColumn("[dodgerblue1]Zip Code[/]").Centered().Width(15).NoWrap());

        if (contacts == null || contacts.Count == 0)
        {
            AnsiConsole.MarkupLine("[red]No records found.[/]");
            return;
        }

        var sortedContacts = contacts.OrderBy(contact => contact.Id).ToList();

        foreach (var contact in sortedContacts)
        {
            table.AddRow(
                contact.Id.ToString(),
                contact.FirstName ?? string.Empty,
                contact.LastName ?? string.Empty,
                contact.PhoneNumber ?? string.Empty,
                contact.EmailAddress ?? string.Empty,
                contact.AddressLine1 ?? string.Empty,
                contact.AddressLine2 ?? string.Empty,
                contact.City ?? string.Empty,
                contact.State ?? string.Empty,
                contact.ZipCode ?? string.Empty);
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
