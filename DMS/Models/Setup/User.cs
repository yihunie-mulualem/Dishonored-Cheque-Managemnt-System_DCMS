using DCMS.Models.Setup;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DCMS.Models.Setup
{
    public class User
    {
        public int Id { get; set; }
        [Required]
        public string FullName { get; set; }
        [Required]
        public string? UserName { get; set; }
        [Required]
        [DataType(DataType.EmailAddress)]
        public string? Email { get; set; }
        [Required]
        [DataType(DataType.Password)]
        [StringLength(20, MinimumLength = 3)]
        public string? Password { get; set; }

        [NotMapped]
        [DataType(DataType.Password)]
        [StringLength(20, MinimumLength = 3)]
        public string? NewPassword { get; set; }
        [NotMapped]
        [DataType(DataType.Password)]
        [StringLength(20, MinimumLength = 3)]
        public string? ConfiermPassword { get; set; }
        public int BranchId { get; set; }
        // R/n
        [Display(Name = "BranchName")]
        //public virtual int BranchId { get; set; }
        public Branch? Branch { get; set; }
        public int RoleId { get; set; }
        [Display(Name = "Role")]
        //public virtual int RoleId { get; set; }
        public Role? Role { get; set; }

        [Required]
        [DefaultValue(true)]
        public bool viewStatus { get; set; } = true;
    }
}
