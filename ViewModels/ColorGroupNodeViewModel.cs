using System;
using System.Collections.Generic;
using System.Linq;
using MorphosPowerPointAddIn.Models;

namespace MorphosPowerPointAddIn.ViewModels
{
    public sealed class ColorGroupNodeViewModel : TreeNodeViewModel
    {
        private readonly string _displayName;
        private readonly IReadOnlyList<ColorNodeViewModel> _items;
        private readonly int _usesCount;
        private bool _childrenLoaded;

        public ColorGroupNodeViewModel(
            string displayName,
            int usesCount,
            IReadOnlyList<ColorNodeViewModel> items,
            bool initiallyExpanded)
        {
            _displayName = displayName;
            _usesCount = usesCount;
            _items = items ?? Array.Empty<ColorNodeViewModel>();
            IsExpanded = initiallyExpanded && _items.Count > 0;
        }

        public override string DisplayName => _displayName;

        public override string UsesText => _usesCount.ToString();

        public override bool CanExpand => _items.Count > 0;

        public string ColorCountText => _items.Count + (_items.Count == 1 ? " color" : " colors");

        protected override void OnExpansionChanged(bool isExpanded)
        {
            if (!isExpanded || _childrenLoaded)
            {
                return;
            }

            Children.ReplaceRange(_items.Cast<TreeNodeViewModel>());

            _childrenLoaded = true;
        }
    }
}
