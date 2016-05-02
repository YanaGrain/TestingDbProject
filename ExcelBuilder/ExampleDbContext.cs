using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBuilder
{
    public class ExampleDbContext : DbContext
    {
        public ExampleDbContext()
            : base("ExampleDbContext")
        {
        }

        public DbSet<ExampleObject> Objects { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ExampleObject>().HasKey(o => o.Id);
            modelBuilder.Entity<ExampleObject>().Property(o => o.Age).IsRequired();
            modelBuilder.Entity<ExampleObject>().Property(o => o.Name).IsRequired();
            modelBuilder.Entity<ExampleObject>().Property(o => o.City).IsRequired();
            modelBuilder.Entity<ExampleObject>().Property(o => o.OnVacation).IsRequired();
            modelBuilder.Entity<ExampleObject>().Property(o => o.Prop1).IsRequired();
            modelBuilder.Entity<ExampleObject>().Property(o => o.Prop2).IsRequired();
            modelBuilder.Entity<ExampleObject>().Property(o => o.Prop3).IsRequired();
            modelBuilder.Entity<ExampleObject>().Property(o => o.Prop4).IsRequired();
            modelBuilder.Entity<ExampleObject>().Property(o => o.Prop5).IsRequired();
            modelBuilder.Entity<ExampleObject>().Property(o => o.Prop6).IsRequired();
        }
    }
}
