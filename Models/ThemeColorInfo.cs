namespace MorphosPowerPointAddIn.Models
{
    public sealed class ThemeColorInfo
    {
        public string SchemeName { get; set; }

        public string DisplayName { get; set; }

        public string HexValue { get; set; }

        public string BrushValue => "#" + (HexValue ?? "000000");
    }
}
