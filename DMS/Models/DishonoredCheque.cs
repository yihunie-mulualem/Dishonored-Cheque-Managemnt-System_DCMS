using DCMS.Models.Setup;
using Microsoft.EntityFrameworkCore.Metadata;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DCMS.Models
{
    public class DishonoredCheque
    {
        public int Id { get; set; }
        public int Branch { get; set; }
        [DataType(DataType.Text)]
        public string HomeBranch { get; set; }
        public string? SubCity { get; set; }
        [DataType(DataType.PhoneNumber)]
        [RegularExpression(@"^\+?(\d[\d-. ]+)?(\([\d-. ]+\))?[\d-. ]+\d$", ErrorMessage = "Invalid phone number.")]
        public string Telephone { get; set; }
        [DataType(DataType.Date)]
        [Display(Name = "Instance Date")]
        public DateTime? InstanceDate { get; set; }
        [Display(Name = "TIN number")]
        [RegularExpression("^[0-9 ]*$", ErrorMessage = "Only numeric characters and spaces are allowed.")]
        public string TIN_Number { get; set; }
        [Display(Name = "Account Number")]
        [MinLength(13)]
        [RegularExpression("^[0-9 ]*$", ErrorMessage = "Only numeric characters and spaces are allowed.")]
        public string AccountNumber { get; set; }
        [Display(Name = "full Name")]
        [DataType(DataType.Text)]
        [RegularExpression("^[a-zA-Z ]*$", ErrorMessage = "Only alphabetic characters and spaces are allowed.")]
        public string FullName { get; set; }
        [DataType(DataType.Date)]
        [Display(Name = "Issue Date")]
        public DateTime IssueDate { get; set; }
        [Display(Name = "Cheque Number")]
        [RegularExpression("^[0-9 ]*$", ErrorMessage = "Only numeric characters and spaces are allowed.")]
        public string ChequeNumber { get; set; }
        [Display(Name = "Beneficiary Name")]
        [DataType(DataType.Text)]
        [RegularExpression("^[a-zA-Z ]*$", ErrorMessage = "Only alphabetic characters and spaces are allowed.")]

        public string BeneficiaryName { get; set; }
        [Column(TypeName = "money")]
        [Display(Name = "AmountOf Money")]
        public Decimal AmountOfMoney { get; set; }
        [Column(TypeName = "money")]
        [Display(Name = "Current Balance")]
        public Decimal CurrentBalance { get; set; }
        public Boolean IsExpired { get; set; }
        [Display(Name = "Expiry Date")]
        public DateTime? ExpiryDate { get; set; }
        [Display(Name = "Registration By")]
        [DataType(DataType.Text)]
        [RegularExpression("^[a-zA-Z ]*$", ErrorMessage = "Only alphabetic characters and spaces are allowed.")]

        public string? RegisteredBy { get; set; }
        [DataType(DataType.Date)]
        [Display(Name = "Registeration Date")]
        public DateTime RegisterationDate { get; set; }
        [Display(Name = "Authorization By")]
        [DataType(DataType.Text)]
        public string? AuthorizedBy { get; set; }
        [DataType(DataType.Date)]
        [Display(Name = "Authorization Date")]
        public DateTime? AuthorizationDate { get; set; }
        [DataType(DataType.Text)]

        public string? Remark { get; set; }
        public Boolean IsBUNotified { get; set; }

        // R/Ship
        public int statusId { get; set; }

        public Status Status { get; set; }
        [Display(Name = "Instance type")]
        public int InstanceTypeId { get; set; }
        public InstanceType? InstanceType { get; set; }

    }
}
