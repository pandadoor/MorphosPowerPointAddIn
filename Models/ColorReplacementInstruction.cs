using System.Collections.Generic;

namespace MorphosPowerPointAddIn.Models
{
    public sealed class ColorReplacementInstruction
    {
        public ColorUsageKind UsageKind { get; set; }

        public string SourceHexValue { get; set; }

        public bool UseThemeColor { get; set; }

        public string ThemeSchemeName { get; set; }

        public string ThemeDisplayName { get; set; }

        public string ReplacementHexValue { get; set; }

        public IReadOnlyList<FontUsageLocation> TargetLocations { get; set; }
    }
}
