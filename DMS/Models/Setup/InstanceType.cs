namespace DCMS.Models.Setup
{
    public class InstanceType
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public ICollection<DishonoredCheque> DishonoredCheques { get; set;}
    }
}
