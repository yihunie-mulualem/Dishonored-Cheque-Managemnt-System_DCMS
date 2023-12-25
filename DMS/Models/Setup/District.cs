namespace DCMS.Models.Setup
{
    public class District
    {
        public int Id { get; set; }
        public string Code { get; set; }
        public string Name { get; set; }
        public ICollection<Branch> Branches { get; set; }
    }
}
