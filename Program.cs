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
        var (excelHeaders, excelTableName) = dataRepository.ExtractHeadersFromExcel(filePath);
        var fileName = dataRepository.CreateTable(filePath, excelHeaders, excelTableName);
        var excelData = dataRepository.ReadExcelData(filePath);
        dataRepository.SeedData(fileName, excelTableName, excelData);
        bool hasIdColumn = dataRepository.CheckIfIdColumnExistsInExcel(filePath);
        Display.PrintAllData(excelTableName, hasIdColumn);
        break;
    case ".csv":
        var (csvHeaders, csvTableName) = dataRepository.ExtractHeadersFromCsv(filePath);
        fileName = dataRepository.CreateTable(filePath, csvHeaders, csvTableName);
        var csvData = dataRepository.ReadCsvData(filePath);
        dataRepository.SeedData(fileName, csvTableName, csvData);
        hasIdColumn = dataRepository.CheckIfIdColumnExistsInCsv(filePath);
        Display.PrintAllData(csvTableName, hasIdColumn);
        break;
    case ".pdf":
        var (pdfFields, pdfTableName) = dataRepository.ExtractFieldNamesFromPdf(filePath);
        fileName = dataRepository.CreateTable(filePath, pdfFields, pdfTableName);
        var pdfData = dataRepository.ReadPdfData(filePath);
        dataRepository.SeedData(fileName, pdfTableName, pdfData);
        Display.PrintAllData(pdfTableName, false);
        break;
    default:
        Console.WriteLine("Unsupported file type.");
        break;
}