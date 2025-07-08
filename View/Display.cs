namespace ExcelReader.View;

using ExcelReader.Model;
using Spectre.Console;

public class Display
{
    public static void PrintAllContacts(string heading)
    {
        var repository = new ContactsRepository(DatabaseUtility.GetConnectionString());
        var contacts = repository.GetAllContacts();

        Console.Clear();

        var rule = new Rule($"[green]{heading}[/]");
        rule.Justification = Justify.Left;
        AnsiConsole.Write(rule);

        var table = new Table()
            .Border(TableBorder.Rounded)
            .AddColumn(new TableColumn("[dodgerblue1]ID[/]").Centered())
            .AddColumn(new TableColumn("[dodgerblue1]First Name[/]").Centered())
            .AddColumn(new TableColumn("[dodgerblue1]Last Name[/]").Centered())
            .AddColumn(new TableColumn("[dodgerblue1]Phone Number[/]").Centered())
            .AddColumn(new TableColumn("[dodgerblue1]Email Address[/]").Centered())
            .AddColumn(new TableColumn("[dodgerblue1]Address Line 1[/]").Centered())
            .AddColumn(new TableColumn("[dodgerblue1]Address Line 2[/]").Centered())
            .AddColumn(new TableColumn("[dodgerblue1]City[/]").Centered())
            .AddColumn(new TableColumn("[dodgerblue1]State[/]").Centered())
            .AddColumn(new TableColumn("[dodgerblue1]Zip Code[/]").Centered());

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

        AnsiConsole.Write(table);
    }
}
