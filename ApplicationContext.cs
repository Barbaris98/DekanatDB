using Microsoft.EntityFrameworkCore;
using Microsoft.VisualBasic.ApplicationServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DekanatDB
{
    public class ApplicationContext : DbContext
    {
        public DbSet<Faculty> Facultys { get; set; } = null!;
        public DbSet<Student> Students { get; set; } = null!;
        public ApplicationContext()
        {
            Database.EnsureCreated();   // создаем бд с новой схемой

        }
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlite("Data Source=Students.db");
        }

    }
}
