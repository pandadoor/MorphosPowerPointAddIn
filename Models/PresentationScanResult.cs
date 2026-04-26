using System;
using System.Collections.Generic;

namespace MorphosPowerPointAddIn.Models
{
    public sealed class PresentationScanResult
    {
        public IReadOnlyList<FontInventoryItem> FontItems { get; set; } = Array.Empty<FontInventoryItem>();

        public ColorScanResult ColorScanResult { get; set; } = new ColorScanResult();
    }
}
