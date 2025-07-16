using DocumentFormat.OpenXml.Spreadsheet;
using ExcelReader.Model;

namespace ExcelReader.Controller;

public class DataController
{
    private readonly DataRepository _dataRepository;

    public DataController(DataRepository dataRepository)
    {
        _dataRepository = dataRepository;
    }

    public void UpdateExcel(string filePath, List<Dictionary<string, object>> data)
    {
        _dataRepository.UpdateExcelData(filePath, data);

        
    }

    public void UpdateCsv(string filePath, Dictionary<string, object> data)
    {
        _dataRepository.UpdateCsvData(filePath, data);
    }

    public void UpdatePdf(string filePath, Dictionary<string, object> data)
    {
        _dataRepository.UpdatePdfData(filePath, data);
    }

    public List<Dictionary<string, object>> AddRowsColumn(List<Dictionary<string, object>> data)
    {
        List<Dictionary<string, object>> dataWithRowsColumn = new List<Dictionary<string, object>>();

        for (int i  = 0; i < data.Count; i++)
        {
            var dataWithRow = 
        }
    }
}
