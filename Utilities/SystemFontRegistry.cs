using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Media;
using Microsoft.Win32;

namespace MorphosPowerPointAddIn.Utilities
{
    internal static class SystemFontRegistry
    {
        private static readonly Regex SuffixPattern = new Regex(@"\s*\(.*\)\s*$", RegexOptions.Compiled);
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
            AddRegistryFonts(Registry.LocalMachine, fonts);
            AddRegistryFonts(Registry.CurrentUser, fonts);

            try
            {
                using (var installed = new InstalledFontCollection())
                {
                    foreach (var family in installed.Families)
                    {
                        fonts.Add(family.Name);
                    }
                }
            }
            catch
            {
            }

            try
            {
                foreach (var family in Fonts.SystemFontFamilies)
                {
                    try
                    {
                        if (!string.IsNullOrWhiteSpace(family.Source))
                        {
                            fonts.Add(family.Source.Trim());
                        }

                        foreach (var localizedName in family.FamilyNames.Values)
                        {
                            if (!string.IsNullOrWhiteSpace(localizedName))
                            {
                                fonts.Add(localizedName.Trim());
                            }
                        }

                        AddTypefaceAliases(family, fonts);
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }

            AddShellFontNames(fonts);

            var snapshot = fonts.ToList();

            lock (SyncRoot)
            {
                if (_cachedFonts == null)
                {
                    _cachedFonts = snapshot;
                }

                return _cachedFonts;
            }
        }

        public static bool SystemFontExists(string fontName)
        {
            var normalized = NormalizeFontName(fontName);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return false;
            }

            return GetInstalledFontNames().Contains(normalized, StringComparer.OrdinalIgnoreCase);
        }

        private static void AddRegistryFonts(RegistryKey root, ISet<string> fonts)
        {
            try
            {
                using (var key = root.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"))
                {
                    if (key == null)
                    {
                        return;
                    }

                    foreach (var valueName in key.GetValueNames())
                    {
                        var normalized = NormalizeRegistryName(valueName);
                        if (!string.IsNullOrWhiteSpace(normalized))
                        {
                            fonts.Add(normalized);
                        }
                    }
                }
            }
            catch
            {
            }
        }

        private static void AddTypefaceAliases(FontFamily family, ISet<string> fonts)
        {
            if (family == null || fonts == null)
            {
                return;
            }

            foreach (var typeface in family.GetTypefaces())
            {
                try
                {
                    GlyphTypeface glyphTypeface;
                    if (!typeface.TryGetGlyphTypeface(out glyphTypeface) || glyphTypeface == null)
                    {
                        continue;
                    }

                    AddFamilyNames(glyphTypeface.FamilyNames, fonts);
                    AddFamilyNames(glyphTypeface.Win32FamilyNames, fonts);

                    AddCompositeNames(glyphTypeface.FamilyNames, glyphTypeface.FaceNames, fonts);
                    AddCompositeNames(glyphTypeface.Win32FamilyNames, glyphTypeface.Win32FaceNames, fonts);
                }
                catch
                {
                }
            }
        }

        private static void AddFamilyNames(
            IDictionary<System.Globalization.CultureInfo, string> names,
            ISet<string> fonts)
        {
            if (names == null || fonts == null)
            {
                return;
            }

            foreach (var familyName in names.Values)
            {
                AddFontName(fonts, familyName);
            }
        }

        private static void AddCompositeNames(
            IDictionary<System.Globalization.CultureInfo, string> familyNames,
            IDictionary<System.Globalization.CultureInfo, string> faceNames,
            ISet<string> fonts)
        {
            if (familyNames == null || faceNames == null || fonts == null)
            {
                return;
            }

            var faceNameList = faceNames.Values
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var familyName in familyNames.Values.Where(x => !string.IsNullOrWhiteSpace(x)))
            {
                AddFontName(fonts, familyName);

                foreach (var faceName in faceNameList)
                {
                    AddFontName(fonts, familyName + " " + faceName);
                }
            }
        }

        private static void AddShellFontNames(ISet<string> fonts)
        {
            object shellApplication = null;
            object shellNamespace = null;
            object shellItems = null;

            try
            {
                var shellType = Type.GetTypeFromProgID("Shell.Application");
                if (shellType == null)
                {
                    return;
                }

                shellApplication = Activator.CreateInstance(shellType);
                if (shellApplication == null)
                {
                    return;
                }

                dynamic shell = shellApplication;
                shellNamespace = shell.Namespace(0x14);
                if (shellNamespace == null)
                {
                    return;
                }

                dynamic fontNamespace = shellNamespace;
                shellItems = fontNamespace.Items();
                if (shellItems == null)
                {
                    return;
                }

                var itemCount = Convert.ToInt32(fontNamespace.Items().Count);
                for (var index = 0; index < itemCount; index++)
                {
                    try
                    {
                        dynamic item = fontNamespace.Items().Item(index);
                        AddFontName(fonts, Convert.ToString(item.Name));
                        ReleaseComObject(item);
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
            finally
            {
                ReleaseComObject(shellItems);
                ReleaseComObject(shellNamespace);
                ReleaseComObject(shellApplication);
            }
        }

        private static void AddFontName(ISet<string> fonts, string fontName)
        {
            var normalized = NormalizeFontName(fontName);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return;
            }

            fonts.Add(normalized);
        }

        private static void ReleaseComObject(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                try
                {
                    Marshal.ReleaseComObject(value);
                }
                catch
                {
                }
            }
        }

        private static string NormalizeRegistryName(string valueName)
        {
            if (string.IsNullOrWhiteSpace(valueName))
            {
                return string.Empty;
            }

            return SuffixPattern.Replace(valueName, string.Empty).Trim();
        }

        private static string NormalizeFontName(string fontName)
        {
            return string.IsNullOrWhiteSpace(fontName) ? string.Empty : fontName.Trim();
        }
    }
}
