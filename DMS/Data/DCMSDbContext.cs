using DCMS.Models;
using DCMS.Models.Setup;
using Microsoft.EntityFrameworkCore;
using System.Linq;

namespace DCMS.Data
{
    public class DCMSDbContext : DbContext
    {
        public DCMSDbContext(DbContextOptions options) : base(options)
        {
        }


        public DbSet<Role> Roles { get; set; }
        public DbSet<User> Users { get; set; }
        public DbSet<BlockedList> BlockedLists { get; set; }
        public DbSet<Branch> Branches  { get; set; }
        public DbSet<District> Districts  { get; set; }
        public DbSet<InstanceType> InstanceTypes  { get; set; }
        public DbSet<DishonoredCheque> DishonoredCheques  { get; set; }
        public DbSet<AuthorizeStatus> AuthorizeStatus { get; set; }


        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            
            modelBuilder.Entity<Role>().HasData
                (
                new Role() { Id = 1, Name = "ADMIN" },
                new Role() { Id = 2, Name = "CHEKER" },
                new Role() { Id = 3, Name = "MAKER" }
                ); 
            modelBuilder.Entity<District>().HasData
                (
                   new District { Id = 1, Code = "01", Name = "HEAD OFFICE" }
                );
            modelBuilder.Entity<Branch>().HasData
                (
                   new Branch { Id = 1, Code = "001", Name = "Bole Brance", DistrictId=1 },
                   new Branch { Id = 2, Code = "002", Name = "HAYAHULET MAZORIA BRANCH",DistrictId=1},
                   new Branch { Id = 3, Code = "999", Name = "HEAD OFFICE", DistrictId=1}
                );
            modelBuilder.Entity<User>().HasData
                (
                new User() { Id = 1, FullName = "ADMIN", UserName = "Admin", Email = "admin@berhanbanksc.com", Password = "MTIz", viewStatus = true,BranchId=1,RoleId=1}
               ); 
        }
    }
}

