using MorphosPowerPointAddIn.Models;

namespace MorphosPowerPointAddIn.ViewModels
{
    public sealed class FontUsageNodeViewModel : TreeNodeViewModel
    {
        public FontUsageNodeViewModel(FontUsageLocation location)
        {
            Location = location;
        }

        public FontUsageLocation Location { get; }

        public override string DisplayName => Location.Label;
    }
}
