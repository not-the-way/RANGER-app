using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RANGER.Database
{
    internal class DatabaseImplementation
    {
        private static readonly string connectionString = @"Data Source= RANGER-database.db";

        public DatabaseImplementation() { }

        public void Query(string sql, SQLiteParameter[] parameters = null)
        {
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                var command = new SQLiteCommand(sql, connection);

                if (parameters != null)
                {
                    foreach (var item in parameters)
                    {
                        command.Parameters.Add(item);
                    }
                }

                command.ExecuteNonQuery();
            }
        }

        public class Employee
        {
            public int Employee_ID { get; set; }
            public string FullName { get; set; }
            public DateTime Date { get; set; }
            public string Login {  get; set; }
            public string Password {  get; set; }
            public string AccessLevel {  get; set; }
        }
    }
}
