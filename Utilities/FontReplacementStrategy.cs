namespace MorphosPowerPointAddIn.Utilities
{
    internal static class FontReplacementStrategy
    {
        public static bool ShouldUsePackageMutation(bool isSaved, string filePath)
        {
            return isSaved && !string.IsNullOrWhiteSpace(FontNameNormalizer.Normalize(filePath));
        }
    }
}
