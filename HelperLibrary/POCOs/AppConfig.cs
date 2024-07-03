using System.Collections.Generic;

namespace HelperLibrary.POCOs
{
    public class AppConfig
    {
        public string AppName { get; set; }
        public string SourceItems { get; set; }
        public bool CopySourceItems { get; set; }
        public string Mode { get; set; }
        public string File { get; set; }
        public string Arguments { get; set; }
        public List<string> TargetComputers { get; set; }
        public int Timeout { get; set; }
        public string DestinationPath { get; set; }
    }
}