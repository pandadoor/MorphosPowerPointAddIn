namespace MorphosPowerPointAddIn.Models
{
    public sealed class FontUsageLocation
    {
        public PresentationScope Scope { get; set; }

        public int? SlideIndex { get; set; }

        public int? ShapeId { get; set; }

        public string ScopeName { get; set; }

        public string ShapeName { get; set; }

        public string Label { get; set; }

        public bool IsSelectable { get; set; }
    }
}
