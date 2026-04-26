using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using MorphosPowerPointAddIn.Models;

namespace MorphosPowerPointAddIn.Services
{
    internal sealed class FontScanSessionCache
    {
        private readonly Func<ISet<string>> _installedFontsProvider;
        private readonly Func<PowerPoint.Presentation, ISet<string>, FontScanSnapshot> _snapshotFactory;
        private readonly Action<FontScanSnapshot, string> _metadataRefresher;
        private string _identity;
        private FontScanSnapshot _snapshot;

        public FontScanSessionCache(
            Func<ISet<string>> installedFontsProvider,
            Func<PowerPoint.Presentation, ISet<string>, FontScanSnapshot> snapshotFactory,
            Action<FontScanSnapshot, string> metadataRefresher)
        {
            _installedFontsProvider = installedFontsProvider ?? throw new ArgumentNullException(nameof(installedFontsProvider));
            _snapshotFactory = snapshotFactory ?? throw new ArgumentNullException(nameof(snapshotFactory));
            _metadataRefresher = metadataRefresher ?? throw new ArgumentNullException(nameof(metadataRefresher));
        }

        public FontScanSnapshot GetOrCreateSnapshot(PowerPoint.Presentation presentation)
        {
            if (presentation == null)
            {
                return new FontScanSnapshot();
            }

            var identity = BuildIdentity(presentation);
            if (_snapshot != null && string.Equals(identity, _identity, StringComparison.OrdinalIgnoreCase))
            {
                return _snapshot;
            }

            var installedFonts = _installedFontsProvider();
            var snapshot = _snapshotFactory(presentation, installedFonts);
            _metadataRefresher(snapshot, snapshot.FilePath);

            _snapshot = snapshot;
            _identity = identity;
            return snapshot;
        }

        public void Invalidate()
        {
            _snapshot = null;
            _identity = null;
        }

        private static string BuildIdentity(PowerPoint.Presentation presentation)
        {
            if (presentation == null)
            {
                return string.Empty;
            }

            var path = ReadPath(presentation);

            return path
                + "|"
                + ReadSaved(presentation)
                + "|"
                + ReadLastWriteTicks(path)
                + "|"
                + ReadCount(() => presentation.Slides.Count)
                + "|"
                + ReadCount(() => presentation.Fonts.Count);
        }

        private static string ReadPath(PowerPoint.Presentation presentation)
        {
            try
            {
                return presentation.FullName ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static int ReadSaved(PowerPoint.Presentation presentation)
        {
            try
            {
                return Convert.ToInt32(presentation.Saved);
            }
            catch
            {
                return (int)MsoTriState.msoFalse;
            }
        }

        private static int ReadCount(Func<int> read)
        {
            try
            {
                return read();
            }
            catch
            {
                return -1;
            }
        }

        private static long ReadLastWriteTicks(string filePath)
        {
            try
            {
                return !string.IsNullOrWhiteSpace(filePath) && File.Exists(filePath)
                    ? File.GetLastWriteTimeUtc(filePath).Ticks
                    : 0L;
            }
            catch
            {
                return 0L;
            }
        }
    }

    internal sealed class FontScanSnapshot
    {
        public FontScanSnapshot()
        {
            Fonts = new List<PresentationFontMetadata>();
            CachedScannedFontNames = Array.Empty<string>();
            EmbeddedFontNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            ThemeFontNames = Array.Empty<string>();
            ReplacementTargets = Array.Empty<FontReplacementTarget>();
        }

        public string FilePath { get; set; }

        public bool IsSaved { get; set; }

        public bool CanUsePackageScan { get; set; }

        public bool HasPackageEmbeddingSettings { get; set; }

        public bool RequestsEmbeddedFonts { get; set; }

        public bool SaveSubsetFonts { get; set; }

        public bool HasEmbeddedFontDataKnown { get; set; }

        public bool HasEmbeddedFontData { get; set; }

        public ISet<string> EmbeddedFontNames { get; set; }

        public IReadOnlyList<string> ThemeFontNames { get; set; }

        public IReadOnlyList<string> CachedScannedFontNames { get; set; }

        public IReadOnlyList<FontInventoryItem> CachedFontItems { get; set; }

        public ColorScanResult CachedColorScanResult { get; set; }

        public PresentationScanResult CachedPresentationScanResult { get; set; }

        public IReadOnlyList<FontReplacementTarget> ReplacementTargets { get; set; }

        public int ReplacementTargetsVersion { get; set; }

        public IList<PresentationFontMetadata> Fonts { get; }
    }

    internal sealed class PresentationFontMetadata
    {
        public string FontName { get; set; }

        public bool IsInstalled { get; set; }

        public bool IsEmbeddable { get; set; }

        public bool HasEmbeddableMetadata { get; set; }

        public bool IsEmbedded { get; set; }

        public bool HasEmbeddedMetadata { get; set; }
    }
}
