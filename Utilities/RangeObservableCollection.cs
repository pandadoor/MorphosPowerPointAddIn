using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;

namespace MorphosPowerPointAddIn.Utilities
{
    public sealed class RangeObservableCollection<T> : ObservableCollection<T>
    {
        private bool _suppressNotifications;

        public void ReplaceRange(IEnumerable<T> items)
        {
            CheckReentrancy();

            _suppressNotifications = true;
            try
            {
                Items.Clear();
                if (items != null)
                {
                    foreach (var item in items)
                    {
                        Items.Add(item);
                    }
                }
            }
            finally
            {
                _suppressNotifications = false;
            }

            OnPropertyChanged(new PropertyChangedEventArgs("Count"));
            OnPropertyChanged(new PropertyChangedEventArgs("Item[]"));
            OnCollectionChanged(new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset));
        }

        protected override void OnCollectionChanged(NotifyCollectionChangedEventArgs e)
        {
            if (_suppressNotifications)
            {
                return;
            }

            base.OnCollectionChanged(e);
        }

        protected override void OnPropertyChanged(PropertyChangedEventArgs e)
        {
            if (_suppressNotifications)
            {
                return;
            }

            base.OnPropertyChanged(e);
        }
    }
}
