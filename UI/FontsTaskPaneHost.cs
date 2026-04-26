using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using MorphosPowerPointAddIn.ViewModels;

namespace MorphosPowerPointAddIn.UI
{
    public sealed class FontsTaskPaneHost : UserControl
    {
        private const int MinimumPaneWidth = 300;
        private const int MinimumPaneHeight = 540;

        public FontsTaskPaneHost(FontsPaneViewModel viewModel)
        {
            Dock = DockStyle.Fill;
            MinimumSize = new Size(MinimumPaneWidth, MinimumPaneHeight);

            var elementHost = new ElementHost
            {
                Dock = DockStyle.Fill,
                MinimumSize = new Size(MinimumPaneWidth, MinimumPaneHeight),
                Child = new FontsUserControl(viewModel)
            };

            Controls.Add(elementHost);
        }
    }
}
