namespace WordsReplace.Models
{
    public class MergeRequest
    {
        public string FilePath1 { get; set; } = string.Empty;
        public string FilePath2 { get; set; } = string.Empty;
        public string MergedFilePath { get; set; } = string.Empty;
    }
}
