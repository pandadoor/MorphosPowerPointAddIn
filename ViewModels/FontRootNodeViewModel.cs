using System.Collections.Generic;
using System.Linq;
using MorphosPowerPointAddIn.Models;

namespace MorphosPowerPointAddIn.ViewModels
{
    public sealed class FontRootNodeViewModel : TreeNodeViewModel
    {
        private readonly IReadOnlyList<FontInventoryItem> _items;

        public FontRootNodeViewModel(IReadOnlyList<FontInventoryItem> items)
        {
            _items = items ?? new List<FontInventoryItem>();
            IsExpanded = true;
        }

        public override string DisplayName => "Fonts (" + _items.Count + ")";

        public override string UsesText => _items.Sum(x => x.UsesCount).ToString();

        public override string EmbeddingText
        {
            get
            {
                if (_items.Count == 0)
                {
                    return "No fonts";
                }

                if (_items.All(x => x.EmbeddingStatus == FontEmbeddingStatus.Yes))
                {
                    return "Embedded";
                }

                if (_items.All(x => x.EmbeddingStatus == FontEmbeddingStatus.Subset))
                {
                    return "Subset";
                }

                if (_items.All(x => x.EmbeddingStatus == FontEmbeddingStatus.No))
                {
                    return "Not embedded";
                }

                return "Mixed";
            }
        }

        public override bool HasStatus => _items.Any(x => x.IsSubstituted || x.HasSaveWarning);

        public override string StatusText
        {
            get
            {
                var substitutedCount = _items.Count(x => x.IsSubstituted);
                var saveWarningCount = _items.Count(x => x.HasSaveWarning);

                if (substitutedCount > 0 && saveWarningCount > 0)
                {
                    return substitutedCount + " substituted, " + saveWarningCount + " save warnings";
                }

                if (substitutedCount > 0)
                {
                    return substitutedCount + " substituted";
                }

                return saveWarningCount > 0 ? saveWarningCount + " can't embed" : string.Empty;
            }
        }

        public override bool HasSaveWarning => _items.Any(x => x.HasSaveWarning);

        public override string StatusToolTip
        {
            get
            {
                if (_items.Any(x => x.IsSubstituted))
                {
                    return "One or more stored fonts differ from the fonts PowerPoint is currently rendering.";
                }

                return HasSaveWarning
                    ? "One or more fonts can still trigger PowerPoint's font-availability warning when embedded save is requested."
                    : string.Empty;
            }
        }
    }
}
