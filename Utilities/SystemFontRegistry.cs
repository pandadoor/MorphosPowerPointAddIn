using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Media;
using System.Text.RegularExpressions;

namespace MorphosPowerPointAddIn.Utilities
{
    internal static class SystemFontRegistry
    {
        private static readonly object SyncRoot = new object();
        private static IReadOnlyCollection<string> _cachedFonts;

        public static IReadOnlyCollection<string> GetInstalledFontNames()
        {
            lock (SyncRoot)
            {
                if (_cachedFonts != null)
                {
                    return _cachedFonts;
                }
            }

            var fonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                // Primary source: WPF SystemFontFamilies (Very fast)
                foreach (var family in Fonts.SystemFontFamilies)
                {
                    foreach (var name in family.FamilyNames.Values)
                    {
                        AddFontName(fonts, name);
                    }
                }

                // SECONARY SOURCE: Windows Registry (Instant, exactly what Windows/Office uses)
                // HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts
                try
                {
                    using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"))
                    {
                        if (key != null)
                        {
                            foreach (var valueName in key.GetValueNames())
                            {
                                // Registry values are often "Font Name (TrueType)" or similar.
                                var name = valueName;
                                var bracketIndex = name.IndexOf('(');
                                if (bracketIndex > 0)
                                {
                                    name = name.Substring(0, bracketIndex).Trim();
                                }
                                AddFontName(fonts, name);
                            }
                        }
                    }
                }
                catch { }

                // FALLBACK: DFS scan of C:\Windows\Fonts for unregistered font files.
                // We only do this for the system fonts folder as requested.
                // We use a high-performance check to avoid re-parsing known fonts.
                ScanFontsDirectoryDFS(@"C:\Windows\Fonts", fonts);
            }
            catch
            {
                // Fail gracefully
            }

            var snapshot = fonts.OrderBy(x => x).ToList();

            lock (SyncRoot)
            {
                if (_cachedFonts == null)
                {
                    _cachedFonts = snapshot;
                }

                return _cachedFonts;
            }
        }

        private static void ScanFontsDirectoryDFS(string directoryPath, ISet<string> fonts)
        {
            try
            {
                if (!System.IO.Directory.Exists(directoryPath)) return;

                var files = System.IO.Directory.GetFiles(directoryPath);
                foreach (var file in files)
                {
                    var ext = System.IO.Path.GetExtension(file);
                    if (string.Equals(ext, ".ttf", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(ext, ".ttc", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(ext, ".otf", StringComparison.OrdinalIgnoreCase))
                    {
                        // OPTIMIZATION: Only parse the file if we haven't found many fonts yet
                        // or if we really need to. Parsing thousands of font files is the cause of UI lag.
                        if (fonts.Count > 1000) continue; 

                        try
                        {
                            foreach (var family in Fonts.GetFontFamilies(new Uri(file, UriKind.Absolute)))
                            {
                                foreach (var name in family.FamilyNames.Values)
                                {
                                    AddFontName(fonts, name);
                                }
                            }
                        }
                        catch { }
                    }
                }

                foreach (var dir in System.IO.Directory.GetDirectories(directoryPath))
                {
                    ScanFontsDirectoryDFS(dir, fonts);
                }
            }
            catch { }
        }

        public static bool SystemFontExists(string fontName)
        {
            if (string.IsNullOrWhiteSpace(fontName)) return false;
            return GetInstalledFontNames().Contains(fontName.Trim(), StringComparer.OrdinalIgnoreCase);
        }

        private static void AddFontName(ISet<string> fonts, string fontName)
        {
            if (string.IsNullOrWhiteSpace(fontName)) return;
            var trimmed = fontName.Trim();
            if (trimmed.Length > 0)
            {
                fonts.Add(trimmed);
            }
        }
    }
}
