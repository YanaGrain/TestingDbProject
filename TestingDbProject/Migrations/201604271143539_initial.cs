namespace TestingDbProject.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class initial : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.FakeObjects",
                c => new
                    {
                        Id = c.Long(nullable: false, identity: true),
                        Name = c.String(nullable: false),
                        Age = c.Int(nullable: false),
                        City = c.String(nullable: false),
                        OnVacation = c.Boolean(nullable: false),
                        Prop1 = c.Int(nullable: false),
                        Prop2 = c.Double(nullable: false),
                        Prop3 = c.Decimal(nullable: false, precision: 18, scale: 2),
                        Prop4 = c.Single(nullable: false),
                        Prop5 = c.String(nullable: false),
                        Prop6 = c.String(nullable: false),
                    })
                .PrimaryKey(t => t.Id);
            
        }
        
        public override void Down()
        {
            DropTable("dbo.FakeObjects");
        }
    }
}
