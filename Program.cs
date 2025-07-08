using ExcelReader.Model;
using Microsoft.Data.SqlClient;

Console.Title = "Excel Reader";

var contactsRepository = new ContactsRepository(DatabaseUtility.GetConnectionString());

contactsRepository.RecreateDatabase();
contactsRepository.CreateTable();

if (DatabaseUtility.CountRows("Stacks") == 0)
    contactsRepository.SeedContacts();

