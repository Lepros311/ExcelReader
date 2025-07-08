using Microsoft.Data.SqlClient;
using System.Data;
using Dapper;

namespace ExcelReader.Model;

public class ContactsRepository
{

    private readonly string _connectionString;
    private readonly string _targetDbName;

    public ContactsRepository(string connectionString)
    {
        _connectionString = connectionString;
        _targetDbName = "ContactsDb";
    }

    public void RecreateDatabase()
    {
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
            Console.WriteLine($"Database {_targetDbName} recreated.");
        }
    }

    public void CreateTable()
    {
        using (SqlConnection connection = new SqlConnection(_connectionString))
        {
            connection.Open();
            string createTableQuery = @"
                    IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Contacts')
                    BEGIN
                        CREATE TABLE Contacts (
                            Id INT IDENTITY(1,1) PRIMARY KEY,
                            FirstName NVARCHAR(100) NOT NULL,
                            LastName NVARCHAR(100) NOT NULL,
                            PhoneNumber NVARCHAR(100) NULL,
                            EmailAddress NVARCHAR(100) NULL,
                            AddressLine1 NVARCHAR(100) NULL,
                            AddressLine2 NVARCHAR(100) NULL,
                            City NVARCHAR(100) NULL,
                            State NVARCHAR(100) NULL,
                            ZipCode NVARCHAR(100) NULL                           
                        );
                    END;";

            using (SqlCommand command = new SqlCommand(createTableQuery, connection))
            {
                try
                {
                    command.ExecuteNonQuery();
                }
                catch (SqlException ex)
                {
                    Console.WriteLine($"An error occurred while creating the Contacts table: {ex.Message}");
                }
            }
        }
    }

    public void SeedContacts()
    {
        using (SqlConnection connection = new SqlConnection(_connectionString))
        {
            connection.Open();
            string insertContactsQuery = "INSERT INTO Contacts (StackName) VALUES (@StackName); SELECT SCOPE_IDENTITY();";

            var stackNames = new List<string> { "Math", "Science", "History" };

            foreach (var stackName in stackNames)
            {
                using (SqlCommand command = new SqlCommand(insertContactsQuery, connection))
                {
                    command.Parameters.AddWithValue("@StackName", stackName);
                    command.ExecuteScalar();
                }
            }
        }
    }

    public List<Contact> GetAllContacts()
    {
        var contacts = new List<Contact>();

        using (SqlConnection connection = new SqlConnection(_connectionString))
        {
            connection.Open();
            string query = @"SELECT * FROM Contacts ORDER BY Id DESC";

            using (var command = new SqlCommand(query, connection))
            using (var reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    var contact = new Contact
                    {
                        Id = reader.GetInt32(0),
                        FirstName = reader.GetString(1),
                        LastName = reader.GetString(2),
                        PhoneNumber = reader.GetString(3),
                        EmailAddress = reader.GetString(4),
                        AddressLine1 = reader.GetString(5),
                        AddressLine2 = reader.GetString(6),
                        City = reader.GetString(7),
                        State = reader.GetString(8),
                        ZipCode = reader.GetString(9),
                    };
                    contacts.Add(contact);
                }
            }
        }

        return contacts;
    }

}
