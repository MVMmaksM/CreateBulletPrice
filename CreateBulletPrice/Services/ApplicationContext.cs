using CreateBulletPrice.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateBulletPrice.Services
{
    internal class ApplicationContext:DbContext
    {
        public DbSet<PerechenModelKor> Perechen_kor { get; set; } 
        public DbSet<PerechenModelPolny> Perechen_polny { get; set; }

        public ApplicationContext()=> Database.EnsureCreated();
            
      
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer("Initial Catalog=Apk_rc_bullet;Data Source=p45-db08;Trusted_Connection=True;TrustServerCertificate=True");
        }

        //protected override void OnModelCreating(ModelBuilder modelBuilder)
        //{
        //    modelBuilder.Entity<PerechenModel>().HasNoKey();
        //}
    }
}
