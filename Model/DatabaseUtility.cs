using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace ExcelReader.Model;

public class DatabaseUtility
{
    public static string GetConnectionString()
    {
        string currentDirectory = Directory.GetCurrentDirectory();

        string projectDirectory = Path.Combine(currentDirectory, @"..\..\..");

        var configuration = new ConfigurationBuilder()
        .SetBasePath(Directory.GetCurrentDirectory())
        .AddJsonFile($"{projectDirectory}\\app.json", optional: true, reloadOnChange: true)
        .Build();

        string? connectionString = configuration.GetConnectionString("connection");

        return connectionString;
    }

    public static int CountRows(string tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ArgumentException("Table name cannot be null or empty.", nameof(tableName));
        }

        var validTables = new[] { "Contacts" };
        if (!Array.Exists(validTables, t => t.Equals(tableName, StringComparison.OrdinalIgnoreCase)))
        {
            throw new ArgumentException("Invalid table name.", nameof(tableName));
        }

        string query = $"SELECT COUNT(1) FROM {tableName};";

        using (var connection = new SqlConnection(GetConnectionString()))
        {
            using (var command = new SqlCommand(query, connection))
            {
                connection.Open();
                return (int)command.ExecuteScalar();
            }
        }
    }
}
