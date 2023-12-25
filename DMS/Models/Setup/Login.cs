using System.ComponentModel.DataAnnotations;

namespace DCMS.Models.Setup
{
    public class Login
    {
        [Required]
        public string UserName { get; set; }
        [Required]
       public string Password { get; set; }
    }
}
