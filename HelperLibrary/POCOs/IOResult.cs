using System.Collections.Generic;

namespace HelperLibrary.POCOs
{
    public class IOResult
    {
        public string PCAddress { get; set; }
        public string SourcePath { get; set; }
        public List<string> SourcePaths { get; set; }
        public string DestinationPath { get; set; }
        public List<string> DestinationPaths { get; set; }
        public bool OperationSuccessful { get; set; }
        public string ErrorMessage { get; set; }
    }
}