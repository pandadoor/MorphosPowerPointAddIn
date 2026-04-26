namespace MorphosPowerPointAddIn.Models
{
    public sealed class ScanProgressInfo
    {
        public int CompletedItems { get; set; }

        public int TotalItems { get; set; }

        public string Message { get; set; }

        public double Percentage => TotalItems == 0 ? 0d : (double)CompletedItems / TotalItems * 100d;
    }
}
