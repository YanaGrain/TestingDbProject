using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Entity;

namespace TestingDbProject
{
    public class FakeDbContext : DbContext
    {
        public FakeDbContext()
            : base("FakeDbContext")
        {
        }

        public DbSet<FakeObject> Objects { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<FakeObject>().HasKey(o=>o.Id);
            modelBuilder.Entity<FakeObject>().Property(o=>o.Age).IsRequired();
            modelBuilder.Entity<FakeObject>().Property(o => o.Name).IsRequired();
            modelBuilder.Entity<FakeObject>().Property(o => o.City).IsRequired();
            modelBuilder.Entity<FakeObject>().Property(o => o.OnVacation).IsRequired();
            modelBuilder.Entity<FakeObject>().Property(o => o.Prop1).IsRequired();
            modelBuilder.Entity<FakeObject>().Property(o => o.Prop2).IsRequired();
            modelBuilder.Entity<FakeObject>().Property(o => o.Prop3).IsRequired();
            modelBuilder.Entity<FakeObject>().Property(o => o.Prop4).IsRequired();
            modelBuilder.Entity<FakeObject>().Property(o => o.Prop5).IsRequired();
            modelBuilder.Entity<FakeObject>().Property(o => o.Prop6).IsRequired();

            //base.OnModelCreating(modelBuilder);
        }
    }
}
