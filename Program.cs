using ExcelReader.Model;
using ExcelReader.View;
using OfficeOpenXml;
using System.IO.Enumeration;

Console.Title = "Excel Reader";

ExcelPackage.License.SetNonCommercialPersonal("Andrew");

var dataRepository = new DataRepository(DatabaseUtility.GetConnectionString());

dataRepository.RecreateDatabase();

var userInterface = new UserInterface();
string filePath = userInterface.GetFilePath();

var (tableName, fileName) = dataRepository.CreateTableFromExcel(filePath);
dataRepository.SeedData(fileName, tableName);

bool hasIdColumn = dataRepository.CheckIfIdColumnExistsInExcel(filePath);
Display.PrintAllData(tableName, hasIdColumn);