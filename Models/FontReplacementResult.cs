using System;
using System.Collections.Generic;
using System.Linq;

namespace MorphosPowerPointAddIn.Models
{
    public sealed class FontReplacementResult
    {
        public FontReplacementResult()
        {
            RemainingSubstitutedFonts = Array.Empty<string>();
            RemainingNonEmbeddableFonts = Array.Empty<string>();
        }

        public IReadOnlyList<string> RemainingSubstitutedFonts { get; set; }

        public IReadOnlyList<string> RemainingNonEmbeddableFonts { get; set; }

        public bool SaveValidationCopySucceeded { get; set; }

        public bool SaveValidationDetectedEmbeddedFontData { get; set; }

        public bool HasWarnings =>
            RemainingSubstitutedFonts.Count > 0
            || RemainingNonEmbeddableFonts.Count > 0;

        public string WarningMessage
        {
            get
            {
                if (!HasWarnings)
                {
                    return string.Empty;
                }

                var sections = new List<string>();
                if (RemainingSubstitutedFonts.Count > 0)
                {
                    sections.Add("Still substituted: " + string.Join(", ", RemainingSubstitutedFonts));
                }

                if (RemainingNonEmbeddableFonts.Count > 0)
                {
                    sections.Add("Cannot embed safely: " + string.Join(", ", RemainingNonEmbeddableFonts));
                }

                if (SaveValidationCopySucceeded
                    && !SaveValidationDetectedEmbeddedFontData
                    && RemainingNonEmbeddableFonts.Count > 0)
                {
                    sections.Add("The embedded validation copy did not contain embedded font data.");
                }

                if (RemainingSubstitutedFonts.Count > 0)
                {
                    sections.Add("These fonts are still being substituted in PowerPoint.");
                }

                if (RemainingNonEmbeddableFonts.Count > 0)
                {
                    sections.Add("These fonts can still trigger PowerPoint's font-availability warning on save.");
                }

                return string.Join(Environment.NewLine + Environment.NewLine, sections.Where(x => !string.IsNullOrWhiteSpace(x)));
            }
        }
    }
}
