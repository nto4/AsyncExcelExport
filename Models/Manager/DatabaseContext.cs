using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace TestNagis.Models.Manager
{
    public class DatabaseContext : DbContext
    {
        public DbSet<Transection> Transections { get; set; }
        public DbSet<Download> Downloads { get; set; }

        public DatabaseContext()
        {
            //Database.SetInitializer<DatabaseContext>(new DBCreator());
            Database.SetInitializer<DatabaseContext>(new DBCreator());
        }
    }
    
    public class DBCreator : CreateDatabaseIfNotExists<DatabaseContext>
    {
        protected override void Seed(DatabaseContext context)
        {
            List<Transection> transectionList = new List<Transection>();
            for (int i = 0; i < 1000; i++)
            {
                Transection islem = new Transection();
                islem.Amount = FakeData.NumberData.GetDouble();
                islem.Buyer = FakeData.NameData.GetFirstName();
                islem.Seller = FakeData.NameData.GetFirstName();
                islem.Date = FakeData.DateTimeData.GetDatetime();
                transectionList.Add(islem);
                context.Transections.Add(islem);
            }

            context.SaveChanges();

            for (int i = 0; i < 1000; i++)
            {
                Download d = new Download();
                d.CreateDate = FakeData.DateTimeData.GetDatetime();
                d.EndDate = FakeData.DateTimeData.GetDatetime();
                d.StartDate = FakeData.DateTimeData.GetDatetime();
                d.IsExist = FakeData.BooleanData.GetBoolean();
                d.GuidName = FakeData.NameData.GetCompanyName();

                context.Downloads.Add(d);
            }

            context.SaveChanges();
        }
    }
}
