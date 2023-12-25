namespace DCMS.Models.Setup
{
    public class BlockedList
    {
        public int Id { get; set; }
        public DishonoredCheque DishonoredCheque { get; set; }
        public Boolean HasExpired { get; set; }
        public DateTime BlockedDate { get; set; }
        public string Remark { get; set; }
    }
}
