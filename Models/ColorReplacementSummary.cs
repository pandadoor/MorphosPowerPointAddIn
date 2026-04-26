namespace MorphosPowerPointAddIn.Models
{
    public sealed class ColorReplacementSummary
    {
        public int ReplacementCount { get; set; }

        public int RemainingDirectColors { get; set; }

        public int RemainingUses { get; set; }

        public bool PreviewAvailable { get; set; }

        public bool Applied { get; set; }
    }
}
