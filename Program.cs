using ExcelReader.Model;
using ExcelReader.View;
using OfficeOpenXml;

Console.Title = "Excel Reader";

ExcelPackage.License.SetNonCommercialPersonal("Andrew");

var dataRepository = new DataRepository(DatabaseUtility.GetConnectionString());

dataRepository.RecreateDatabase();

var userInterface = new UserInterface();
string filePath = userInterface.GetFilePath();

string extension = Path.GetExtension(filePath);

switch (extension)
{
    case ".xlsx":
        var (headers, tableName) = dataRepository.ExtractHeadersFromExcel(filePath);
        break;
    case ".csv":
        var (headers, tableName) = dataRepository.ExtractHeadersFromCsv(filePath);
        break;
    default:
        Console.WriteLine("Unsupported file type.");
        break;
}

var fileName = dataRepository.CreateTable(filePath, headers, tableName);
dataRepository.SeedData(fileName, tableName);

bool hasIdColumn = dataRepository.CheckIfIdColumnExistsInExcel(filePath);
Display.PrintAllData(tableName, hasIdColumn);