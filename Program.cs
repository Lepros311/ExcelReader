using ExcelReader.Model;
using ExcelReader.View;
using OfficeOpenXml;

Console.Title = "Excel Reader";

ExcelPackage.License.SetNonCommercialPersonal("Andrew");

var contactsRepository = new ContactsRepository(DatabaseUtility.GetConnectionString());

contactsRepository.RecreateDatabase();
contactsRepository.CreateTable();

if (DatabaseUtility.CountRows("Contacts") == 0)
    contactsRepository.SeedContacts();

Display.PrintAllContacts("Contacts");