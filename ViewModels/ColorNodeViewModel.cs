using System;
using System.Collections.Generic;
using System.Linq;
using MorphosPowerPointAddIn.Models;

namespace MorphosPowerPointAddIn.ViewModels
{
    public sealed class ColorNodeViewModel : TreeNodeViewModel
    {
        private ColorInventoryItem _item;
        private IReadOnlyList<FontUsageLocation> _locations;
        private bool _childrenLoaded;

        public ColorNodeViewModel(ColorInventoryItem item)
        {
            UpdateItem(item);
        }

        public ColorInventoryItem Item => _item;

        public string HexText => "#" + (Item.HexValue ?? "000000");

        public string ThemeMatchText => Item.MatchesThemeColor
            ? Item.MatchingThemeDisplayName
            : "Direct RGB";

        public bool HasThemeMatch => Item.MatchesThemeColor;

        public override string DisplayName => HexText;

        public override string UsesText => Item.UsesCount.ToString();

        public override bool CanExpand => _locations.Count > 0;

        protected override void OnExpansionChanged(bool isExpanded)
        {
            if (!isExpanded)
            {
                return;
            }

            EnsureChildrenLoaded();
        }

        internal void UpdateItem(ColorInventoryItem item)
        {
            _item = item ?? new ColorInventoryItem
            {
                HexValue = "000000",
                RgbValue = "0, 0, 0",
                Locations = Array.Empty<FontUsageLocation>()
            };
            _locations = _item.Locations ?? Array.Empty<FontUsageLocation>();

            if (_childrenLoaded)
            {
                Children.ReplaceRange(_locations.Select(location => (TreeNodeViewModel)new ColorUsageNodeViewModel(location)));
            }

            OnPropertyChanged(nameof(Item));
            OnPropertyChanged(nameof(HexText));
            OnPropertyChanged(nameof(ThemeMatchText));
            OnPropertyChanged(nameof(HasThemeMatch));
            OnPropertyChanged(nameof(DisplayName));
            OnPropertyChanged(nameof(UsesText));
            OnPropertyChanged(nameof(CanExpand));
        }

        private void EnsureChildrenLoaded()
        {
            if (_childrenLoaded)
            {
                return;
            }

            Children.ReplaceRange(_locations.Select(location => (TreeNodeViewModel)new ColorUsageNodeViewModel(location)));
            _childrenLoaded = true;
        }
    }
}
