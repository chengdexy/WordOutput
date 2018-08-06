using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.ModelConfiguration.Conventions;
using System.Linq;
using System.Text;
using WordOutput.Models;

namespace WordOutput.DAL {
    class MainContext : DbContext {
        public MainContext() : base("MainContext") { }

        //DbSets
        public DbSet<DocModel> DocModels { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder) {
            modelBuilder.Conventions.Remove<PluralizingTableNameConvention>();
            base.OnModelCreating(modelBuilder);
        }
    }
}
