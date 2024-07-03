namespace HelperLibrary.POCOs
{
    public class DHCPScope
    {
        public string ScopeId { get; set; }
        public string SubnetMask { get; set; }
        public string Name { get; set; }
        public string State { get; set; }
        public string StartRange { get; set; }
        public string EndRange { get; set; }
        public string LeaseDuration { get; set; }
    }
}