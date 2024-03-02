namespace WordsReplace.Models
{
    public class ReplaceWordRequest
    {
        public string PathDocDefault { get; set; } = string.Empty;
        public string PathNewDoc { get; set; } = string.Empty;
        public string OldWord { get; set; } = string.Empty;
        public string NewWord { get; set; } = string.Empty;
    }
}
