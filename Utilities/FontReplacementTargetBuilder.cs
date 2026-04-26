using System;
using System.Collections.Generic;
using System.Linq;
using MorphosPowerPointAddIn.Models;

namespace MorphosPowerPointAddIn.Utilities
{
    internal static class FontReplacementTargetBuilder
    {
        public static IReadOnlyList<FontReplacementTarget> Build(
            IEnumerable<string> powerPointFontNames,
            IEnumerable<string> systemFontNames,
            IEnumerable<string> sourceFontNames,
            IEnumerable<string> themeFontNames = null)
        {
            var sourceLookup = new HashSet<string>(
                (sourceFontNames ?? Array.Empty<string>())
                    .Select(FontNameNormalizer.Normalize)
                    .Where(x => !string.IsNullOrWhiteSpace(x)),
                StringComparer.OrdinalIgnoreCase);

            var installedFontLookup = new HashSet<string>(
                (systemFontNames ?? Array.Empty<string>())
                    .Select(FontNameNormalizer.NormalizeReplacementFont)
                    .Where(x => !string.IsNullOrWhiteSpace(x)),
                StringComparer.OrdinalIgnoreCase);

            var results = new List<FontReplacementTarget>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var sortKey = 0;

            AddTargets(results, seen, sourceLookup, powerPointFontNames, installedFontLookup, false, true, ref sortKey);
            AddTargets(results, seen, sourceLookup, themeFontNames, installedFontLookup, true, false, ref sortKey);
            AddTargets(results, seen, sourceLookup, installedFontLookup, installedFontLookup, false, false, ref sortKey);

            return results;
        }

        private static void AddTargets(
            ICollection<FontReplacementTarget> results,
            ISet<string> seen,
            ISet<string> sourceLookup,
            IEnumerable<string> fontNames,
            ISet<string> installedFontLookup,
            bool isThemeFont,
            bool isPresentationFont,
            ref int sortKey)
        {
            if (fontNames == null)
            {
                return;
            }

            foreach (var fontName in fontNames)
            {
                var normalizedName = FontNameNormalizer.NormalizeReplacementFont(fontName);
                if (string.IsNullOrWhiteSpace(normalizedName)
                    || sourceLookup.Contains(normalizedName)
                    || !seen.Add(normalizedName))
                {
                    continue;
                }

                var isInstalled = isThemeFont
                    || (installedFontLookup != null && installedFontLookup.Contains(normalizedName));
                if (!isInstalled)
                {
                    continue;
                }

                results.Add(new FontReplacementTarget
                {
                    DisplayName = isThemeFont ? BuildThemeDisplayName(normalizedName) : normalizedName,
                    NormalizedName = normalizedName,
                    IsThemeFont = isThemeFont,
                    IsInstalled = isInstalled,
                    IsPresentationFont = isPresentationFont,
                    SortKey = sortKey++
                });
            }
        }

        private static string BuildThemeDisplayName(string normalizedName)
        {
            if (string.Equals(normalizedName, "+mj-lt", StringComparison.OrdinalIgnoreCase))
            {
                return "+mj-lt (Theme Headings)";
            }

            if (string.Equals(normalizedName, "+mn-lt", StringComparison.OrdinalIgnoreCase))
            {
                return "+mn-lt (Theme Body)";
            }

            return normalizedName;
        }
    }
}
