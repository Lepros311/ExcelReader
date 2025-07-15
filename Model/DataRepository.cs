using CsvHelper;
using Dapper;
using iText.Forms;
using iText.Kernel.Pdf;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using System.Data;
using System.Globalization;
using System.Text;

namespace ExcelReader.Model;

public class DataRepository
{

    private readonly string _connectionString;
    private readonly string _targetDbName;

    public DataRepository(string connectionString)
    {
        _connectionString = connectionString;
        _targetDbName = "ExcelReaderDb";
    }

    public void RecreateDatabase()
    {
        Console.WriteLine("Creating database...");

        var masterConnectionString = $"{_connectionString};Database=master;";

        using (IDbConnection db = new SqlConnection(masterConnectionString))
        {
            db.Open();

            string dropDbSql = $@"
            IF EXISTS (SELECT name FROM sys.databases WHERE name = '{_targetDbName}')
            BEGIN
                ALTER DATABASE [{_targetDbName}] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
                DROP DATABASE [{_targetDbName}];
            END";
            db.Execute(dropDbSql);

            db.Execute($"CREATE DATABASE [{_targetDbName}]");
            Console.WriteLine($"Database {_targetDbName} created.\n");
        }
    }

    public (List<string> headers, string tableName) ExtractHeadersFromExcel(string filePath)
    {
        string tableName = Path.GetFileNameWithoutExtension(filePath); // Get the file name without extension

        var headers = new List<string>();

        using (SqlConnection connection = new SqlConnection(_connectionString))
        {
            connection.Open();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.Columns;

                // Read the header row
                for (int col = 1; col <= colCount; col++)
                {
                    headers.Add(worksheet.Cells[1, col].Text);
                }
            }
        }

        return (headers, tableName);
    }

    public (List<string> headers, string tableName) ExtractFieldNamesFromPdf(string filePath)
    {
        string tableName = Path.GetFileNameWithoutExtension(filePath);
        var fieldNames = new List<string>();

        using (PdfReader pdfReader = new PdfReader(filePath))
        using (PdfDocument pdfDocument = new PdfDocument(pdfReader))
        {
            PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDocument, true);
            var fields = form.GetAllFormFields();

            foreach (var field in fields)
            {
                fieldNames.Add(field.Key);
            }
        }

        return (fieldNames, tableName);
    }

    public (List<string> headers, string tableName) ExtractHeadersFromCsv(string filePath)
    {
        string tableName = Path.GetFileNameWithoutExtension(filePath);

        var headers = new List<string>();

        using (var reader = new StreamReader(filePath))
        using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
        {
            csv.Read();
            csv.ReadHeader();
            headers = csv.Context.Reader.HeaderRecord.ToList();
        }

        return (headers, tableName);
    }

    public string CreateTable(string filePath, List<string> headers, string tableName)
    {
        string fileNameWithExtension = Path.GetFileName(filePath);

        Console.WriteLine($"Creating [{tableName}] table...");

        using (SqlConnection connection = new SqlConnection(_connectionString))
        {
            connection.Open();

            // Create the SQL table dynamically
            var createTableQuery = $"CREATE TABLE [{tableName}] (";

            bool hasIdColumn = headers.Any(header => string.Equals(header, "id", StringComparison.OrdinalIgnoreCase));

            if (hasIdColumn)
            {
                createTableQuery += "[Id] INT PRIMARY KEY, ";
            }
            else
            {
                createTableQuery += "Id INT IDENTITY(1,1) PRIMARY KEY, ";
            }

            foreach (var header in headers)
            {
                if (!string.Equals(header, "id", StringComparison.OrdinalIgnoreCase))
                {
                    createTableQuery += $"[{header}] NVARCHAR(MAX), "; // Use NVARCHAR(MAX) for flexibility
                }
            }

            createTableQuery = createTableQuery.TrimEnd(',', ' ') + ");"; // Remove the last comma and close the statement

            // Execute the create table query
            using (var command = new SqlCommand(createTableQuery, connection))
            {
                command.ExecuteNonQuery();
                Console.WriteLine($"Table [{tableName}] created successfully.\n");
            }
        }

        return fileNameWithExtension; // Return the table name
    }

    public List<Dictionary<string, object>> ReadExcelData(string filePath)
    {
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.Rows;
            int colCount = worksheet.Dimension.Columns;

            var data = new List<Dictionary<string, object>>();
            var (headers, _) = ExtractHeadersFromExcel(filePath);

            for (int row = 2; row <= rowCount; row++)
            {
                var rowData = new Dictionary<string, object>();
                for (int col = 1; col <= colCount; col++)
                {
                    string header = headers[col - 1];
                    object value = worksheet.Cells[row, col].Text;
                    rowData[header] = value;
                }
                data.Add(rowData);
            }
            return data;
        }
    }

    public List<Dictionary<string, object>> ReadPdfData(string filePath)
    {
        var data = new List<Dictionary<string, object>>();

        using (PdfReader pdfReader = new PdfReader(filePath))
        using (PdfDocument pdfDocument = new PdfDocument(pdfReader))
        {
            PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDocument, true);
            var fields = form.GetAllFormFields();

            var rowData = new Dictionary<string, object>();

            foreach (var field in fields)
            {
                string fieldName = field.Key;
                object fieldValue = field.Value.GetValue();

                if (fieldValue is PdfString pdfString)
                {
                    rowData[fieldName] = pdfString.GetValue();
                }
                else
                {
                    rowData[fieldName] = fieldValue;
                }
            }

            data.Add(rowData);
        }

        return data;
    }

    public List<Dictionary<string, object>> ReadCsvData(string filePath)
    {
        using (var reader = new StreamReader(filePath))
        using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
        {
            var records = csv.GetRecords<dynamic>().ToList();
            var data = new List<Dictionary<string, object>>();

            foreach (var record in records)
            {
                var rowData = new Dictionary<string, object>();
                foreach (var kvp in (IDictionary<string, object>)record)
                {
                    rowData[kvp.Key] = kvp.Value;
                }
                data.Add(rowData);
            }
            return data;
        }
    }

    public void SeedData(string fileName, string tableName, List<Dictionary<string, object>> data)
    {
        Console.WriteLine($"Populating data from [{fileName}] to [{tableName}] table...");

        using (SqlConnection connection = new SqlConnection(_connectionString))
        {
            connection.Open();

            foreach (var row in data)
            {
                var parameters = new Dictionary<string, object>();
                foreach (var kvp in row)
                {
                    string cleanParameterName = SanitizeColumnName(kvp.Key);
                    parameters[cleanParameterName] = kvp.Value;
                }

                var insertColumns = string.Join(", ", row.Keys.Select(h => $"[{h}]"));
                var insertValues = string.Join(", ", row.Keys.Select(h => $"@{SanitizeColumnName(h)}"));

                string insertQuery = $"INSERT INTO [{tableName}] ({insertColumns}) VALUES ({insertValues})";

                connection.Execute(insertQuery, parameters);
            }
        }

        Console.WriteLine($"Data populated to [{tableName}] table.\n");
    }

    private string SanitizeColumnName(string columnName)
    {
        // Use a StringBuilder to construct the sanitized name
        var sanitized = new StringBuilder();

        // Loop through each character in the column name
        foreach (char c in columnName)
        {
            // Check if the character is a letter, digit, or underscore
            if (char.IsLetterOrDigit(c) || c == '_')
            {
                sanitized.Append(c);
            }
            else if (c == ' ')
            {
                // Replace spaces with an underscore
                sanitized.Append('_');
            }
        }

        // Ensure the name does not start with a digit
        if (sanitized.Length > 0 && char.IsDigit(sanitized[0]))
        {
            sanitized.Insert(0, '_'); // Prepend an underscore if it starts with a digit
        }

        return sanitized.ToString();
    }


    public List<Dictionary<string, object>> GetAllData(string tableName)
    {
        Console.WriteLine("Gathering data for display...\n");

        var data = new List<Dictionary<string, object>>();

        using (SqlConnection connection = new SqlConnection(_connectionString))
        {
            connection.Open();
            string query = $"SELECT * FROM [{tableName}] ORDER BY Id ASC";

            using (var command = new SqlCommand(query, connection))
            using (var reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    var row = new Dictionary<string, object>();
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        string columnName = reader.GetName(i);
                        object value = reader.IsDBNull(i) ? null : reader.GetValue(i);
                        row[columnName] = value;
                    }

                    data.Add(row);
                }
            }
        }

        return data;
    }

    public bool CheckIfIdColumnExistsInExcel(string filePath)
    {
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];

            var headerRow = worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column];

            foreach (var cell in headerRow)
            {
                if (cell.Text.Equals("Id", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
        }
        return false;
    }

    public bool CheckIfIdColumnExistsInCsv(string filePath)
    {
        using (var reader = new StreamReader(filePath))
        using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
        {
            csv.Read();
            csv.ReadHeader();
            var headerRecord = csv.Context.Reader.HeaderRecord;

            foreach (var header in headerRecord)
            {
                if (header.Equals("Id", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
        }
        return false;
    }

}