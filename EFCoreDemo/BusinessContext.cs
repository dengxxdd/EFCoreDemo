using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
namespace EFCoreDemo
{
    public class BusinessContext : DbContext
    {
        public DbSet<Leader> Leaders { get; set; }
        public DbSet<FamilyMember> FamilyMembers { get; set; }

        public DbSet<ReportBusiness> ReportBusinesses { get; set; }
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlite("Data Source=gszz.db");
        }
    }
}
