using System.Collections.Generic;

namespace MorphosPowerPointAddIn.Models
{
    public sealed class ColorInventoryItem
    {
        public ColorUsageKind UsageKind { get; set; }

        public string UsageKindLabel { get; set; }

        public string HexValue { get; set; }

        public string BrushValue => "#" + (HexValue ?? "000000");

        public string RgbValue { get; set; }

        public int UsesCount { get; set; }

        public bool MatchesThemeColor { get; set; }

        public string MatchingThemeDisplayName { get; set; }

        public string MatchingThemeSchemeName { get; set; }

        public IReadOnlyList<FontUsageLocation> Locations { get; set; }
    }
}
