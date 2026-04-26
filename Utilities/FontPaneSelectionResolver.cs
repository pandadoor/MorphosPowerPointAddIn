using System.Collections.Generic;
using System.Linq;

namespace MorphosPowerPointAddIn.Utilities
{
    internal static class FontPaneSelectionResolver
    {
        public static string ResolvePostReplaceSelectionKey(
            string replacementFontName,
            IEnumerable<string> availableFontNames,
            string currentSelectionKey = null)
        {
            var fonts = (availableFontNames ?? System.Array.Empty<string>())
                .Select(FontNameNormalizer.Normalize)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            var replacementKey = BuildFontSelectionKey(replacementFontName);
            if (!string.IsNullOrWhiteSpace(replacementKey)
                && fonts.Any(x => replacementKey == BuildFontSelectionKey(x)))
            {
                return replacementKey;
            }

            if (!string.IsNullOrWhiteSpace(currentSelectionKey)
                && fonts.Any(x => currentSelectionKey == BuildFontSelectionKey(x)))
            {
                return currentSelectionKey;
            }

            return fonts.Count == 0 ? string.Empty : BuildFontSelectionKey(fonts[0]);
        }

        public static string BuildFontSelectionKey(string fontName)
        {
            var normalized = FontNameNormalizer.Normalize(fontName);
            return string.IsNullOrWhiteSpace(normalized) ? string.Empty : "font|" + normalized;
        }
    }
}
