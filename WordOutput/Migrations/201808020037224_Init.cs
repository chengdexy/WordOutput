namespace WordOutput.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class Init : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.DocModel",
                c => new
                    {
                        ID = c.Int(nullable: false, identity: true),
                        Name = c.String(),
                        Gender = c.String(),
                        Age = c.String(),
                        SickYears = c.String(),
                        PhoneNumbers = c.String(),
                        Address = c.String(),
                        VipNumber = c.String(),
                        KindOfSick = c.String(),
                        UsingDrugs = c.String(),
                        CurSymptoms = c.String(),
                        Records = c.String(),
                    })
                .PrimaryKey(t => t.ID);
            
        }
        
        public override void Down()
        {
            DropTable("dbo.DocModel");
        }
    }
}
