using Group4333.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Group4333.Database
{
    public class DatabaseHelper
    {
        string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\zarin\OneDrive\Desktop\kit_kai\isrpo\isrpo2\4333Group-26\Group4333\ServicesDB.mdf;Integrated Security=True";

        public void CreateTable()
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                string createTableQuery = @"
                    IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Services' AND xtype='U')
                    CREATE TABLE Services (
                        Id INT PRIMARY KEY,
                        Name NVARCHAR(200),
                        Type NVARCHAR(100),
                        Price DECIMAL(18,2)
                    )";

                SqlCommand cmd = new SqlCommand(createTableQuery, conn);
                cmd.ExecuteNonQuery();
            }
        }

        public void SaveServices(List<Service> services)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                SqlCommand clearCmd = new SqlCommand("DELETE FROM Services", conn);
                clearCmd.ExecuteNonQuery();

                foreach (var service in services)
                {
                    string insertQuery = @"
                        INSERT INTO Services (Id, Name, Type, Price) 
                        VALUES (@Id, @Name, @Type, @Price)";

                    SqlCommand cmd = new SqlCommand(insertQuery, conn);
                    cmd.Parameters.AddWithValue("@Id", service.Id);
                    cmd.Parameters.AddWithValue("@Name", service.Name);
                    cmd.Parameters.AddWithValue("@Type", service.Type);
                    cmd.Parameters.AddWithValue("@Price", service.Price);

                    cmd.ExecuteNonQuery();
                }
            }
        }
        public List<Service> GetAllServices()
        {
            List<Service> services = new List<Service>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                string selectQuery = "SELECT * FROM Services ORDER BY Id";

                SqlCommand cmd = new SqlCommand(selectQuery, conn);
                SqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    services.Add(new Service
                    {
                        Id = Convert.ToInt32(reader["Id"]),
                        Name = reader["Name"].ToString(),
                        Type = reader["Type"].ToString(),
                        Price = Convert.ToDecimal(reader["Price"])
                    });
                }
            }

            return services;
        }
    }
}
