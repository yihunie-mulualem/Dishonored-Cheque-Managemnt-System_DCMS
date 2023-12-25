namespace DCMS.Models.Setup
{
    public class Role
    {
        public int Id { get; set; }
        public string Name { get; set; }
        // R/n
        public ICollection<User> Users { get; set; }
    }
}
