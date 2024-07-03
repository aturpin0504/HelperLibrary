namespace HelperLibrary.POCOs
{
    public class QUserSession
    {
        public string ComputerName { get; set; }
        public string UserName { get; set; }
        public string SessionName { get; set; }
        public string Id { get; set; }
        public string State { get; set; }
        public string IdleTime { get; set; }
        public string LogonTime { get; set; }
    }
}