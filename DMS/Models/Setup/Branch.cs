using DCMS.Models.Setup;

namespace DCMS.Models.Setup
{
    public class Branch
    {
        public int Id { get; set; }
        public string Code { get; set; }
        public string Name { get; set; }
        public int DistrictId { get; set; }
        public District District { get; set; }
        public ICollection<User> Users { get; set; }
    }
}
