using Dapper;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using System.Data;
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

    public (string tableName, string fileName) CreateTableFromExcel(string filePath)
    {
        string tableName = Path.GetFileNameWithoutExtension(filePath); // Get the file name without extension
        string fileNameWithExtension = Path.GetFileName(filePath);

        Console.WriteLine($"Creating [{tableName}] table...");

        using (SqlConnection connection = new SqlConnection(_connectionString))
        {
            connection.Open();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.Columns;

                // Read the header row
                var headers = new List<string>();
                for (int col = 1; col <= colCount; col++)
                {
                    headers.Add(worksheet.Cells[1, col].Text);
                }

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
        }

        return (tableName, fileNameWithExtension); // Return the table name
    }


    public void SeedData(string fileName, string tableName)
    {
        string projectRoot = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\.."));
        string filePath = Path.Combine(projectRoot, fileName);

        Console.WriteLine($"Populating data from [{fileName}] to [{tableName}] table...");

        using (SqlConnection connection = new SqlConnection(_connectionString))
        {
            connection.Open();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;
                int columnCount = worksheet.Dimension.Columns;

                var headers = new List<string>();
                for (int column = 1; column <= columnCount; column++)
                {
                    string header = worksheet.Cells[1, column].Text.Trim();
                    headers.Add(header);
                }

                for (int row = 2; row <= rowCount; row++)
                {
                    var parameters = new Dictionary<string, object>();
                    for (int column = 1; column <= columnCount; column++)
                    {
                        string header = headers[column - 1];
                        object value = worksheet.Cells[row, column].Text;

                        string cleanParameterName = SanitizeColumnName(header);
                        parameters[cleanParameterName] = value;
                    }

                    var insertColumns = string.Join(", ", headers.Select(h => $"[{h}]"));
                    var insertValues = string.Join(", ", headers.Select(h => $"@{SanitizeColumnName(h)}"));

                    string insertQuery = $"INSERT INTO [{tableName}] ({insertColumns}) VALUES ({insertValues})";

                    connection.Execute(insertQuery, parameters);
                }
            }
        }

        Console.WriteLine($"Data populated from [{fileName}] to [{tableName}] table.\n");
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
        // Ensure the Excel package is disposed properly
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            // Get the first worksheet
            var worksheet = package.Workbook.Worksheets[0];
            // Get the first row (header row)
            var headerRow = worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column];
            // Check if "Id" exists in the header row
            foreach (var cell in headerRow)
            {
                if (cell.Text.Equals("Id", StringComparison.OrdinalIgnoreCase))
                {
                    return true; // "Id" column exists
                }
            }
        }
        return false; // "Id" column does not exist

    }
}