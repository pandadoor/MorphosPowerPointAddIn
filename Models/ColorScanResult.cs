using System;
using System.Collections.Generic;

namespace MorphosPowerPointAddIn.Models
{
    public sealed class ColorScanResult
    {
        public ColorScanResult()
        {
            Items = Array.Empty<ColorInventoryItem>();
            ThemeColors = Array.Empty<ThemeColorInfo>();
        }

        public IReadOnlyList<ColorInventoryItem> Items { get; set; }

        public IReadOnlyList<ThemeColorInfo> ThemeColors { get; set; }
    }
}
