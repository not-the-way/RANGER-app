using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Reflection;
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

        public List<T> ExecuteQuery<T>(string sql, SQLiteParameter[] parameters = null) where T : new()
        {
            var result = new List<T>();

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();

                using (var command = new SQLiteCommand(sql, connection))
                {
                    if (parameters != null)
                    {
                        foreach (var item in parameters)
                        {
                            command.Parameters.Add(item);
                        }
                    }

                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            T obj = new T();

                            for (int i =  0; i < reader.FieldCount; i++)
                            {
                                string columnName = reader.GetName(i);
                                object value = reader.GetValue(i);

                                var property = typeof(T).GetProperty(columnName, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                                
                                if (property != null && value != DBNull.Value)
                                {
                                    property.SetValue(obj, Convert.ChangeType(value, property.PropertyType));
                                }
                            }

                            foreach (var property in typeof(T).GetProperties())
                            {
                                if (property.PropertyType.IsClass && property.PropertyType != typeof(string))
                                {
                                    var relatedEntity = Activator.CreateInstance(property.PropertyType);
                                    bool hasRelatedData = false;

                                    foreach (var relatedProperty in property.PropertyType.GetProperties())
                                    {
                                        string relatedColumnName = $"{property.Name}_{relatedProperty.Name}";
                                        int relatedOrdinal;
                                        try
                                        {
                                            relatedOrdinal = reader.GetOrdinal(relatedColumnName);
                                        }
                                        catch
                                        {
                                            continue;
                                        }

                                        if (relatedOrdinal >= 0 && reader[relatedColumnName] != DBNull.Value)
                                        {
                                            relatedProperty.SetValue(relatedEntity, Convert.ChangeType(reader[relatedColumnName], relatedProperty.PropertyType));
                                            hasRelatedData = true;
                                        }
                                    }

                                    if (hasRelatedData)
                                    {
                                        property.SetValue(obj, relatedEntity);
                                    }
                                }
                            }
                            result.Add(obj);
                        }
                    }
                }
            }

            return result;
        }

        public class Employee
        {
            public int Employee_ID { get; set; }
            public string FullName { get; set; }
            public DateTime Date { get; set; }
            public string Login { get; set; }
            public string Password { get; set; }
            public string AccessLevel { get; set; }
        }

        class Client
        {
            public int Client_ID { get; set; }
            public string Fullname { get; set; }
            public string Passport { get; set; }
            public bool IsMature { get; set; }
            public bool ReleaseOfLiability { get; set; }
        }

        class Firearm
        {
            public string FirearmSerialNumber { get; set; }
            public string Name { get; set; }
            public string Category { get; set; }
            public string Condition { get; set; }
            public DateTime LastMaitenanceDate { get; set; }
        }

        class FirearmCategory
        {
            public string Category { get; set; }
        }

        class FirearmCondition
        {
            public string Condition { get; set; }
        }

        class Issue
        {
            public int Issue_ID { get; set; }
            public int Client_ID { get; set; }
            public int Employee_ID { get; set; }
            public string FirearmSerialNumber { get; set; }
            public int Item_ID { get; set; }
            public DateTime DateTimeOfIssue {  get; set; }
            public DateTime DateTimeOfReturn { get; set; }
        }

        class Warehouse
        {
            public int Item_ID { get; set; }
            public string ItemName { get; set; }
            public DateTime DeliveryDate { get; set; }
            public string WarehouseType { get; set; }
            public int Quantity { get; set; }
        }
        
        class WarehouseType
        {
            public string Type { get; set; }
        }
    }
}
