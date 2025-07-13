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
        var (ExcelHeaders, ExcelTableName) = dataRepository.ExtractHeadersFromExcel(filePath);
        var fileName = dataRepository.CreateTable(filePath, ExcelHeaders, ExcelTableName);
        var ExcelData = dataRepository.ReadExcelData(filePath);
        dataRepository.SeedData(fileName, ExcelTableName, ExcelData);
        bool hasIdColumn = dataRepository.CheckIfIdColumnExistsInExcel(filePath);
        Display.PrintAllData(ExcelTableName, hasIdColumn);
        break;
    case ".csv":
        var (CsvHeaders, CsvTableName) = dataRepository.ExtractHeadersFromCsv(filePath);
        fileName = dataRepository.CreateTable(filePath, CsvHeaders, CsvTableName);
        var CsvData = dataRepository.ReadCsvData(filePath);
        dataRepository.SeedData(fileName, CsvTableName, CsvData);
        hasIdColumn = dataRepository.CheckIfIdColumnExistsInCsv(filePath);
        Display.PrintAllData(CsvTableName, hasIdColumn);
        break;
    default:
        Console.WriteLine("Unsupported file type.");
        break;
}