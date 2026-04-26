using System;
using System.Collections.Generic;
using System.Linq;
using MorphosPowerPointAddIn.Models;

namespace MorphosPowerPointAddIn.Utilities
{
    internal sealed class FontSearchIndex
    {
        private readonly string _searchText;

        private FontSearchIndex(string searchText)
        {
            _searchText = searchText ?? string.Empty;
        }

        public static FontSearchIndex Create(FontInventoryItem item)
        {
            if (item == null)
            {
                return new FontSearchIndex(string.Empty);
            }

            var tokens = new List<string>();
            AddToken(tokens, item.FontName);

            if (item.Locations != null)
            {
                foreach (var location in item.Locations)
                {
                    AddToken(tokens, location == null ? string.Empty : location.Label);
                }
            }

            return new FontSearchIndex(string.Join("\n", tokens.Distinct(StringComparer.OrdinalIgnoreCase)));
        }

        public bool Matches(string query)
        {
            var normalized = FontNameNormalizer.Normalize(query);
            return string.IsNullOrWhiteSpace(normalized)
                || _searchText.IndexOf(normalized, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void AddToken(ICollection<string> tokens, string value)
        {
            var normalized = FontNameNormalizer.Normalize(value);
            if (!string.IsNullOrWhiteSpace(normalized))
            {
                tokens.Add(normalized);
            }
        }
    }
}
