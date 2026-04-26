using System;
using System.Collections.Generic;
using System.Linq;
using MorphosPowerPointAddIn.Models;
using MorphosPowerPointAddIn.Utilities;

namespace MorphosPowerPointAddIn.ViewModels
{
    public sealed class FontNodeViewModel : TreeNodeViewModel
    {
        private FontInventoryItem _item;
        private IReadOnlyList<FontUsageLocation> _locations;
        private FontSearchIndex _searchIndex;
        private bool _childrenLoaded;

        public FontNodeViewModel(FontInventoryItem item)
        {
            UpdateItem(item);
        }

        public FontInventoryItem Item => _item;

        internal FontSearchIndex SearchIndex => _searchIndex;

        public string FontName => Item.FontName;

        public override string DisplayName => Item.FontName;

        public override string UsesText => Item.UsesCount.ToString();

        public override string EmbeddingText
        {
            get
            {
                if (Item.IsThemeFont)
                {
                    return "Theme";
                }

                switch (Item.EmbeddingStatus)
                {
                    case FontEmbeddingStatus.Yes:
                        return "Embedded";
                    case FontEmbeddingStatus.Subset:
                        return "Subset";
                    case FontEmbeddingStatus.No:
                        return "Not embedded";
                    default:
                        return "Check";
                }
            }
        }

        public string SourceText
        {
            get
            {
                if (Item.IsThemeFont)
                {
                    return "Theme";
                }

                if (Item.IsEmbedded && !Item.IsInstalled)
                {
                    return "Embedded";
                }

                return Item.IsInstalled ? "Installed" : "Missing";
            }
        }

        public override bool IsMissingOrSubstituted => Item.IsMissingOrSubstituted;

        public override bool HasStatus => Item.IsSubstituted || Item.HasSaveWarning;

        public override bool HasSaveWarning => Item.HasSaveWarning;

        public override bool CanExpand => _locations.Count > 0;

        public override string StatusText
        {
            get
            {
                if (Item.IsSubstituted)
                {
                    return "Substituted";
                }

                return Item.HasSaveWarning ? "Warning" : string.Empty;
            }
        }

        public override string StatusToolTip
        {
            get
            {
                if (Item.IsSubstituted && Item.HasSaveWarning && !Item.IsEmbeddable)
                {
                    return "PowerPoint is rendering this text with a different font than the one stored in the file, and the current font still cannot be embedded cleanly.";
                }

                if (Item.IsSubstituted)
                {
                    return "PowerPoint is rendering this text with a different font than the one stored in the file.";
                }

                if (Item.HasSaveWarning)
                {
                    if (Item.IsLocallyMissing)
                    {
                        return "This font is not available on this computer and PowerPoint could not complete a clean embedded validation save.";
                    }

                    return "PowerPoint could not complete a clean embedded validation save for this font.";
                }

                return string.Empty;
            }
        }

        protected override void OnExpansionChanged(bool isExpanded)
        {
            if (!isExpanded)
            {
                return;
            }

            EnsureChildrenLoaded();
        }

        internal void UpdateItem(FontInventoryItem item)
        {
            _item = item ?? new FontInventoryItem
            {
                FontName = string.Empty,
                Locations = Array.Empty<FontUsageLocation>()
            };
            _locations = _item.Locations ?? Array.Empty<FontUsageLocation>();
            _searchIndex = FontSearchIndex.Create(_item);

            if (_childrenLoaded)
            {
                Children.ReplaceRange(_locations.Select(location => (TreeNodeViewModel)new FontUsageNodeViewModel(location)));
            }

            OnPropertyChanged(nameof(Item));
            OnPropertyChanged(nameof(FontName));
            OnPropertyChanged(nameof(DisplayName));
            OnPropertyChanged(nameof(UsesText));
            OnPropertyChanged(nameof(EmbeddingText));
            OnPropertyChanged(nameof(SourceText));
            OnPropertyChanged(nameof(IsMissingOrSubstituted));
            OnPropertyChanged(nameof(HasStatus));
            OnPropertyChanged(nameof(HasSaveWarning));
            OnPropertyChanged(nameof(CanExpand));
            OnPropertyChanged(nameof(StatusText));
            OnPropertyChanged(nameof(StatusToolTip));
        }

        private void EnsureChildrenLoaded()
        {
            if (_childrenLoaded)
            {
                return;
            }

            Children.ReplaceRange(_locations.Select(location => (TreeNodeViewModel)new FontUsageNodeViewModel(location)));
            _childrenLoaded = true;
        }
    }
}
