using System.Collections.Generic;

namespace MorphosPowerPointAddIn.Models
{
    public sealed class FontInventoryItem
    {
        public string FontName { get; set; }

        public int UsesCount { get; set; }

        public FontEmbeddingStatus EmbeddingStatus { get; set; }

        public bool IsInstalled { get; set; }

        public bool IsEmbeddable { get; set; }

        public bool HasEmbeddableMetadata { get; set; }

        public bool HasPresentationMetadata { get; set; }

        public bool IsSubstituted { get; set; }

        public bool HasSaveWarning { get; set; }

        public bool IsThemeFont { get; set; }

        public bool IsEmbedded =>
            EmbeddingStatus == FontEmbeddingStatus.Yes
            || EmbeddingStatus == FontEmbeddingStatus.Subset;

        public bool IsLocallyMissing =>
            !IsThemeFont
            && !IsInstalled
            && !IsEmbedded;

        public bool IsMissingOrSubstituted => IsSubstituted;

        public IReadOnlyList<FontUsageLocation> Locations { get; set; }
    }
}
