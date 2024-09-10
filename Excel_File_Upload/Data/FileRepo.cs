using Excel_File_Upload.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace Excel_File_Upload.Data
{
    public class FileRepo
    {
        private SqlConnection _connection;
        public FileRepo()
        {
            string connectionString = "Data Source=.;Initial Catalog=Excel_File_Upload;Integrated Security=True; TrustServerCertificate = True";


            _connection = new SqlConnection(connectionString);
        }

        // using stored procedure InsertPersonData
        public bool SaveFileInDb(HttpPostedFileBase file)
        {
            try
            {
                // Set the license context for non-commercial use
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(file.InputStream))
                {

                    // Get the first worksheet
                    var worksheet = package.Workbook.Worksheets[0];

                    _connection.Open();

                    // Iterate through rows, assuming the first row is headers
                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        // Extract data from each row
                        string firstName = worksheet.Cells[row, 2].Text;
                        string lastName = worksheet.Cells[row, 3].Text;
                        string gender = worksheet.Cells[row, 4].Text;
                        string country = worksheet.Cells[row, 5].Text;
                        // Check and parse age safely
                        string ageText = worksheet.Cells[row, 6].Text;


                        // Call the InsertPersonData stored procedure
                        using (SqlCommand cmd = new SqlCommand("InsertPersonData", _connection))
                        {
                            cmd.CommandType = System.Data.CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@FirstName", firstName);
                            cmd.Parameters.AddWithValue("@LastName", lastName);
                            cmd.Parameters.AddWithValue("@Gender", gender);
                            cmd.Parameters.AddWithValue("@Country", country);

                            cmd.ExecuteNonQuery();
                        }
                    }

                    _connection.Close();
                    return true;
                }
            }
            catch (Exception ex)
            {
                _connection.Close();
                return false;
            }
        }



        // using sp to get data from db above saved one  and return to home controllers GetFileData method
        public List<Person> GetFileData()
        {
            List<Person> persons = new List<Person>();
            try
            {
                _connection.Open();
                using (SqlCommand cmd = new SqlCommand("GetPersonData", _connection))
                {
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Person person = new Person
                            {
                                Id = Convert.ToInt32(reader["Id"]),
                                FirstName = reader["FirstName"].ToString(),
                                LastName = reader["LastName"].ToString(),
                                Gender = reader["Gender"].ToString(),
                                Country = reader["Country"].ToString(),
                            };

                            persons.Add(person);
                        }
                    }
                }
                _connection.Close();
            }
            catch (Exception ex)
            {
                _connection.Close();
                Console.WriteLine(ex.Message);
            }
            return persons;
        }

    }
}





// stored procedures

// table 
//CREATE TABLE Person (
//    Id INT IDENTITY(1,1) PRIMARY KEY,
//    FirstName VARCHAR(50) NOT NULL,
//    LastName VARCHAR(50) NOT NULL,
//    Gender VARCHAR(20) NOT NULL,
//    Country VARCHAR(50) NOT NULL,
//    Age INT NOT NULL,
//    Date VARCHAR(50) NOT NULL
//);


//GetPersonData
//CREATE PROCEDURE GetPersonData
//AS
//BEGIN
//    SELECT * FROM Person;
//END


//InsertPersonData
//    Create PROCEDURE InsertPersonData
//    @FirstName VARCHAR(50),
//    @LastName VARCHAR(50),
//    @Gender VARCHAR(20),
//    @Country VARCHAR(50)
//AS
//BEGIN
//    INSERT INTO Person (FirstName, LastName, Gender, Country)
//    VALUES (@FirstName, @LastName, @Gender, @Country);
//END
