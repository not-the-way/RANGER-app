using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace RANGER.Database
{
    class DatabaseImplementation
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
            public Employee()
            {
                this.Employees_ID = new HashSet<Issue>();
            }

            public int Employee_ID { get; set; }
            public string FullName { get; set; }
            public DateTime Date { get; set; }
            public string Login { get; set; }
            public string Password { get; set; }
            public string AccessLevel { get; set; }

            public virtual ICollection<Issue> Employees_ID { get; set; }
        }

        public class Client
        {
            public Client()
            {
                this.Clients_ID = new HashSet<Issue>();
            }

            public int Client_ID { get; set; }
            public string Fullname { get; set; }
            public string Passport { get; set; }
            public bool IsMature { get; set; }
            public bool ReleaseOfLiability { get; set; }

            public virtual ICollection<Issue> Clients_ID { get; set; }
        }

        public class Firearm
        {
            public Firearm()
            {
                this.FirearmSerialNumber_ID = new HashSet<Issue>();
            }

            public string FirearmSerialNumber { get; set; }
            public string Name { get; set; }
            public string Category { get; set; }
            public string Condition { get; set; }
            public DateTime LastMaitenanceDate { get; set; }

            public virtual ICollection<Issue> FirearmSerialNumber_ID { get; set; }

            public virtual FirearmCategory FirearmCategory { get; set; }
            public virtual FirearmCondition FirearmCondition { get; set; }
        }

        public class FirearmCategory
        {
            public FirearmCategory()
            {
                this.Category_ID = new HashSet<Firearm>();
            }

            public string Category { get; set; }

            public virtual ICollection<Firearm> Category_ID { get; set; }
        }

        public class FirearmCondition
        {
            public FirearmCondition()
            {
                this.Condition_ID = new HashSet<Firearm>();
            }

            public string Condition { get; set; }

            public virtual ICollection<Firearm> Condition_ID { get; set; }
        }

        public class Issue
        {
            public int Issue_ID { get; set; }
            public int Client_ID { get; set; }
            public int Employee_ID { get; set; }
            public string FirearmSerialNumber { get; set; }
            public int Item_ID { get; set; }
            public DateTime DateTimeOfIssue {  get; set; }
            public DateTime DateTimeOfReturn { get; set; }

            public virtual Client Client { get; set; }

            public virtual Firearm Firearm { get; set; }

            public virtual Employee Employee { get; set; }

            public virtual Warehouse Warehouse { get; set; }
        }

        public class Warehouse
        {
            public Warehouse()
            {
                this.Items_ID = new HashSet<Issue>();
            }

            public int Item_ID { get; set; }
            public string ItemName { get; set; }
            public DateTime DeliveryDate { get; set; }
            public string WarehouseType { get; set; }
            public int Quantity { get; set; }

            public virtual ICollection<Issue> Items_ID { get; set; }

            public virtual WarehouseType WarehouseType_ID { get; set; }
        }
        
        public class WarehouseType
        {
            public WarehouseType()
            {
                this.Type_ID = new HashSet<Warehouse>();
            }

            public string Type { get; set; }

            public virtual ICollection<Warehouse> Type_ID { get; set; }
        }
    }
}
