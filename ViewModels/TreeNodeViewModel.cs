using MorphosPowerPointAddIn.Utilities;

namespace MorphosPowerPointAddIn.ViewModels
{
    public abstract class TreeNodeViewModel : BindableBase
    {
        private bool _isExpanded;

        protected TreeNodeViewModel()
        {
            Children = new RangeObservableCollection<TreeNodeViewModel>();
        }

        public RangeObservableCollection<TreeNodeViewModel> Children { get; }

        public bool IsExpanded
        {
            get => _isExpanded;
            set
            {
                if (SetProperty(ref _isExpanded, value))
                {
                    OnExpansionChanged(value);
                }
            }
        }

        public abstract string DisplayName { get; }

        public virtual bool CanExpand => Children.Count > 0;

        public virtual string UsesText => string.Empty;

        public virtual string EmbeddingText => string.Empty;

        public virtual string StatusText => string.Empty;

        public virtual string StatusToolTip => string.Empty;

        public virtual bool HasStatus => false;

        public virtual bool HasSaveWarning => false;

        public virtual bool IsMissingOrSubstituted => false;

        protected virtual void OnExpansionChanged(bool isExpanded)
        {
        }
    }
}
