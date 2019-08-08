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

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Leader>()
                .Property(b => b.CreateTime)
                .HasDefaultValueSql("datetime('now','localtime')");
            modelBuilder.Entity<FamilyMember>()
                .Property(b => b.CreateTime)
                .HasDefaultValueSql("datetime('now','localtime')");
            modelBuilder.Entity<ReportBusiness>()
                .Property(b => b.CreateTime)
                .HasDefaultValueSql("datetime('now','localtime')");
        }
    }
}
