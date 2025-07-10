using ExcelReader.Model;
using ExcelReader.View;
using OfficeOpenXml;
using System.IO.Enumeration;

Console.Title = "Excel Reader";

ExcelPackage.License.SetNonCommercialPersonal("Andrew");

var contactsRepository = new DataRepository(DatabaseUtility.GetConnectionString());

contactsRepository.RecreateDatabase();

var userInterface = new UserInterface();
string filePath = userInterface.GetFilePath();

var (tableName, fileName) = contactsRepository.CreateTableFromExcel(filePath);
contactsRepository.SeedData(fileName, tableName);

Display.PrintAllData(tableName);