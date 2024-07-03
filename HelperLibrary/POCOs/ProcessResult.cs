namespace HelperLibrary.POCOs
{
    public class ProcessResult
    {
        public string PCAddress { get; set; }
        public int ExitCode { get; set; }
        public string StandardOutput { get; set; }
        public string StandardError { get; set; }

    }
}