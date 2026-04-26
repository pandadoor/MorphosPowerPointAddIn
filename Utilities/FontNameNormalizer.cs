using System;

namespace MorphosPowerPointAddIn.Utilities
{
    internal static class FontNameNormalizer
    {
        public static string Normalize(string fontName)
        {
            return string.IsNullOrWhiteSpace(fontName) ? string.Empty : fontName.Trim();
        }

        public static string NormalizeReplacementFont(string fontName)
        {
            var normalized = Normalize(fontName);
            if (normalized.StartsWith("+mj-lt", StringComparison.OrdinalIgnoreCase))
            {
                return "+mj-lt";
            }

            if (normalized.StartsWith("+mn-lt", StringComparison.OrdinalIgnoreCase))
            {
                return "+mn-lt";
            }

            return normalized;
        }
    }
}
