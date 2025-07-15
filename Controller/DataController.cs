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
}
