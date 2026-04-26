namespace MorphosPowerPointAddIn.Models
{
    public sealed class FontReplacementTarget
    {
        public string DisplayName { get; set; }

        public string NormalizedName { get; set; }

        public bool IsThemeFont { get; set; }

        public bool IsInstalled { get; set; }

        public bool IsPresentationFont { get; set; }

        public int SortKey { get; set; }
    }
}
