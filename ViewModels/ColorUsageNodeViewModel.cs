using MorphosPowerPointAddIn.Models;

namespace MorphosPowerPointAddIn.ViewModels
{
    public sealed class ColorUsageNodeViewModel : TreeNodeViewModel
    {
        public ColorUsageNodeViewModel(FontUsageLocation location)
        {
            Location = location;
        }

        public FontUsageLocation Location { get; }

        public override string DisplayName => Location == null ? string.Empty : Location.Label;
    }
}
