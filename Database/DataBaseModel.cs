using System;
using System.Collections.Generic;

namespace RANGER.Database
{
    public class DataBaseModel
    {
        public class Employee
        {
            public Employee()
            {
                this.Employees_ID = new HashSet<Issue>();
            }

            public int Employee_ID { get; set; }
            public string FullName { get; set; }
            public DateTime EmploymentDate { get; set; }
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
            public DateTime DateTimeOfIssue { get; set; }
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

