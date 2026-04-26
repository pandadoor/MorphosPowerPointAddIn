using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using MorphosPowerPointAddIn.Models;

namespace MorphosPowerPointAddIn.Utilities
{
    internal static class PresentationPackageInspector
    {
        private const string PresentationEntryPath = "ppt/presentation.xml";
        private const string ThemeEntryPath = "ppt/theme/theme1.xml";
        private const string FontsFolderPrefix = "ppt/fonts/";
        private static readonly Regex SlidePattern = new Regex(@"^ppt/slides/slide(\d+)\.xml$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex SlideMasterPattern = new Regex(@"^ppt/slideMasters/slideMaster(\d+)\.xml$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex SlideLayoutPattern = new Regex(@"^ppt/slideLayouts/slideLayout(\d+)\.xml$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex NotesMasterPattern = new Regex(@"^ppt/notesMasters/notesMaster(\d+)\.xml$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly XNamespace PresentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
        private static readonly XNamespace DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";

        public static bool TryGetFontEmbeddingSettings(string filePath, out bool embedFonts, out bool saveSubsetFonts)
        {
            embedFonts = false;
            saveSubsetFonts = false;

            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return false;
            }

            try
            {
                using (var archive = ZipFile.OpenRead(filePath))
                {
                    var entry = archive.Entries.FirstOrDefault(x => x.FullName.Equals(PresentationEntryPath, StringComparison.OrdinalIgnoreCase));
                    if (entry == null)
                    {
                        return false;
                    }

                    using (var stream = entry.Open())
                    {
                        using (var reader = CreateReader(stream))
                        {
                            if (reader.MoveToContent() != XmlNodeType.Element)
                            {
                                return false;
                            }

                            var hasEmbedFonts = TryParseOfficeBooleanAttribute(reader.GetAttribute("embedTrueTypeFonts"), out embedFonts);
                            var hasSaveSubsetFonts = TryParseOfficeBooleanAttribute(reader.GetAttribute("saveSubsetFonts"), out saveSubsetFonts);
                            return hasEmbedFonts || hasSaveSubsetFonts;
                        }
                    }
                }
            }
            catch
            {
                return false;
            }
        }

        public static bool TryGetSaveSubsetFonts(string filePath, out bool saveSubsetFonts)
        {
            bool embedFonts;
            return TryGetFontEmbeddingSettings(filePath, out embedFonts, out saveSubsetFonts);
        }

        public static bool TrySetFontEmbeddingFlags(string filePath, bool embedFonts, bool saveSubsetFonts)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return false;
            }

            try
            {
                using (var archive = ZipFile.Open(filePath, ZipArchiveMode.Update))
                {
                    var entry = archive.Entries.FirstOrDefault(x => x.FullName.Equals(PresentationEntryPath, StringComparison.OrdinalIgnoreCase));
                    if (entry == null)
                    {
                        return false;
                    }

                    XDocument document;
                    using (var readStream = entry.Open())
                    {
                        document = XDocument.Load(readStream);
                    }

                    var root = document.Root;
                    if (root == null)
                    {
                        return false;
                    }

                    root.SetAttributeValue("embedTrueTypeFonts", embedFonts ? "1" : "0");
                    root.SetAttributeValue("saveSubsetFonts", saveSubsetFonts ? "1" : "0");

                    entry.Delete();
                    var replacement = archive.CreateEntry(PresentationEntryPath);
                    using (var writeStream = replacement.Open())
                    {
                        document.Save(writeStream);
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        public static bool TryHasEmbeddedFontData(string filePath, out bool hasEmbeddedFontData)
        {
            hasEmbeddedFontData = false;

            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return false;
            }

            try
            {
                using (var archive = ZipFile.OpenRead(filePath))
                {
                    ISet<string> embeddedFontNames;
                    if (TryGetEmbeddedFontNames(archive, out embeddedFontNames))
                    {
                        hasEmbeddedFontData = embeddedFontNames.Count > 0
                            || archive.Entries.Any(entry => entry.FullName.StartsWith(FontsFolderPrefix, StringComparison.OrdinalIgnoreCase));
                        return true;
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        public static bool TryGetEmbeddedFontNames(string filePath, out ISet<string> embeddedFontNames)
        {
            embeddedFontNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return false;
            }

            try
            {
                using (var archive = ZipFile.OpenRead(filePath))
                {
                    return TryGetEmbeddedFontNames(archive, out embeddedFontNames);
                }
            }
            catch
            {
                embeddedFontNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                return false;
            }
        }

        public static bool TryGetThemeFontNames(string filePath, out IReadOnlyList<string> themeFontNames)
        {
            themeFontNames = Array.Empty<string>();

            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return false;
            }

            try
            {
                using (var archive = ZipFile.OpenRead(filePath))
                {
                    var entry = archive.Entries.FirstOrDefault(x => x.FullName.Equals(ThemeEntryPath, StringComparison.OrdinalIgnoreCase));
                    if (entry == null)
                    {
                        return false;
                    }

                    var tokens = new List<string>(2);
                    using (var stream = entry.Open())
                    {
                        using (var reader = CreateReader(stream))
                        {
                            var frameStack = new Stack<XmlScanFrame>();
                            var currentScheme = string.Empty;

                            while (reader.Read())
                            {
                                if (reader.NodeType == XmlNodeType.Element)
                                {
                                    frameStack.Push(new XmlScanFrame(reader.NamespaceURI, reader.LocalName, false));

                                    if (reader.NamespaceURI == DrawingNamespace.NamespaceName)
                                    {
                                        if (reader.LocalName.Equals("majorFont", StringComparison.OrdinalIgnoreCase))
                                        {
                                            currentScheme = "+mj-lt";
                                        }
                                        else if (reader.LocalName.Equals("minorFont", StringComparison.OrdinalIgnoreCase))
                                        {
                                            currentScheme = "+mn-lt";
                                        }
                                        else if (!string.IsNullOrWhiteSpace(currentScheme)
                                            && reader.LocalName.Equals("latin", StringComparison.OrdinalIgnoreCase))
                                        {
                                            var typeface = NormalizeFontName(reader.GetAttribute("typeface"));
                                            if (!string.IsNullOrWhiteSpace(typeface) && !tokens.Contains(currentScheme, StringComparer.OrdinalIgnoreCase))
                                            {
                                                tokens.Add(currentScheme);
                                            }
                                        }
                                    }

                                    if (reader.IsEmptyElement)
                                    {
                                        PopThemeFontFrame(frameStack, ref currentScheme);
                                    }
                                }
                                else if (reader.NodeType == XmlNodeType.EndElement)
                                {
                                    PopThemeFontFrame(frameStack, ref currentScheme);
                                }
                            }
                        }
                    }

                    themeFontNames = tokens;
                    return true;
                }
            }
            catch
            {
                themeFontNames = Array.Empty<string>();
                return false;
            }
        }

        public static IReadOnlyList<PackageFontUsageRecord> ReadFontUsages(
            string filePath,
            IProgress<ScanProgressInfo> progress,
            CancellationToken cancellationToken)
        {
            try
            {
                return new Services.OpenXmlScanService()
                    .ReadFontUsages(filePath, progress, cancellationToken);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch
            {
                return null;
            }
        }

        private static void ScanShapeContainer(
            IEnumerable<XElement> elements,
            PackageScanEntry scanEntry,
            ICollection<PackageFontUsageRecord> results)
        {
            foreach (var element in elements)
            {
                if (element == null || element.Name.Namespace != PresentationNamespace)
                {
                    continue;
                }

                if (element.Name == PresentationNamespace + "grpSp")
                {
                    ScanShapeContainer(element.Elements(), scanEntry, results);
                    continue;
                }

                if (!IsShapeElement(element))
                {
                    continue;
                }

                var shapeName = ResolveShapeName(element);
                var shapeId = ResolveShapeId(element);
                var fonts = ExtractFontNames(element);
                foreach (var fontName in fonts)
                {
                    results.Add(new PackageFontUsageRecord
                    {
                        FontName = fontName,
                        Location = BuildLocation(scanEntry, shapeName, shapeId)
                    });
                }
            }
        }

        private static FontUsageLocation BuildLocation(PackageScanEntry scanEntry, string shapeName, int? shapeId)
        {
            var label = scanEntry.Scope == PresentationScope.Slide && scanEntry.Index.HasValue
                ? "Slide " + scanEntry.Index.Value + " - " + shapeName
                : scanEntry.Label + " - " + shapeName;

            return new FontUsageLocation
            {
                Scope = scanEntry.Scope,
                SlideIndex = scanEntry.Scope == PresentationScope.Slide ? scanEntry.Index : null,
                ShapeId = shapeId,
                ScopeName = scanEntry.Label,
                ShapeName = shapeName,
                Label = label,
                IsSelectable = scanEntry.Scope == PresentationScope.Slide && scanEntry.Index.HasValue
            };
        }

        private static bool IsShapeElement(XElement element)
        {
            return element.Name == PresentationNamespace + "sp"
                || element.Name == PresentationNamespace + "cxnSp"
                || element.Name == PresentationNamespace + "graphicFrame"
                || element.Name == PresentationNamespace + "pic";
        }

        private static ISet<string> ExtractFontNames(XElement element)
        {
            var fonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var descendant in element.Descendants())
            {
                if (descendant.Name.Namespace != DrawingNamespace)
                {
                    continue;
                }

                var localName = descendant.Name.LocalName;
                if (!localName.Equals("latin", StringComparison.OrdinalIgnoreCase)
                    && !localName.Equals("ea", StringComparison.OrdinalIgnoreCase)
                    && !localName.Equals("cs", StringComparison.OrdinalIgnoreCase)
                    && !localName.Equals("sym", StringComparison.OrdinalIgnoreCase)
                    && !localName.Equals("buFont", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var typeface = NormalizeFontName((string)descendant.Attribute("typeface"));
                if (!string.IsNullOrWhiteSpace(typeface))
                {
                    fonts.Add(typeface);
                }
            }

            return fonts;
        }

        private static string ResolveShapeName(XElement element)
        {
            var properties = element.Descendants(PresentationNamespace + "cNvPr").FirstOrDefault();
            if (properties == null)
            {
                return "Shape";
            }

            var name = ((string)properties.Attribute("name") ?? string.Empty).Trim();
            if (!string.IsNullOrWhiteSpace(name))
            {
                return name;
            }

            var id = ResolveShapeId(element);
            return id.HasValue ? "Shape " + id.Value : "Shape";
        }

        private static int? ResolveShapeId(XElement element)
        {
            var properties = element.Descendants(PresentationNamespace + "cNvPr").FirstOrDefault();
            if (properties == null)
            {
                return null;
            }

            var rawId = (string)properties.Attribute("id");
            int parsedId;
            return int.TryParse(rawId, out parsedId) ? parsedId : (int?)null;
        }

        private static string NormalizeFontName(string fontName)
        {
            return string.IsNullOrWhiteSpace(fontName) ? string.Empty : fontName.Trim();
        }

        private static bool TryGetEmbeddedFontNames(ZipArchive archive, out ISet<string> embeddedFontNames)
        {
            embeddedFontNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (archive == null)
            {
                return false;
            }

            var entry = archive.Entries.FirstOrDefault(x => x.FullName.Equals(PresentationEntryPath, StringComparison.OrdinalIgnoreCase));
            if (entry == null)
            {
                return false;
            }

            try
            {
                using (var stream = entry.Open())
                {
                    using (var reader = CreateReader(stream))
                    {
                        var frameStack = new Stack<XmlScanFrame>();
                        var embeddedFontListDepth = 0;

                        while (reader.Read())
                        {
                            if (reader.NodeType == XmlNodeType.Element)
                            {
                                frameStack.Push(new XmlScanFrame(reader.NamespaceURI, reader.LocalName, false));

                                if (reader.NamespaceURI == PresentationNamespace.NamespaceName
                                    && reader.LocalName.Equals("embeddedFontLst", StringComparison.OrdinalIgnoreCase))
                                {
                                    embeddedFontListDepth = frameStack.Count;
                                }
                                else if (embeddedFontListDepth > 0)
                                {
                                    var fontName = NormalizeFontName(reader.GetAttribute("typeface"));
                                    if (!string.IsNullOrWhiteSpace(fontName))
                                    {
                                        embeddedFontNames.Add(fontName);
                                    }
                                }

                                if (reader.IsEmptyElement)
                                {
                                    PopEmbeddedFontFrame(frameStack, ref embeddedFontListDepth);
                                }
                            }
                            else if (reader.NodeType == XmlNodeType.EndElement)
                            {
                                PopEmbeddedFontFrame(frameStack, ref embeddedFontListDepth);
                            }
                        }

                        return true;
                    }
                }
            }
            catch
            {
                embeddedFontNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                return false;
            }
        }

        private static PackageScanEntry CreateScanEntry(ZipArchiveEntry entry)
        {
            if (entry == null)
            {
                return null;
            }

            var fullName = entry.FullName;
            var slideMatch = SlidePattern.Match(fullName);
            if (slideMatch.Success)
            {
                return new PackageScanEntry(entry, PresentationScope.Slide, ParseIndex(slideMatch), "Slide " + slideMatch.Groups[1].Value, 2000);
            }

            var masterMatch = SlideMasterPattern.Match(fullName);
            if (masterMatch.Success)
            {
                var index = ParseIndex(masterMatch);
                var label = index == 1 ? "Slide master" : "Slide master " + index;
                return new PackageScanEntry(entry, PresentationScope.SlideMaster, null, label, 0);
            }

            var layoutMatch = SlideLayoutPattern.Match(fullName);
            if (layoutMatch.Success)
            {
                var index = ParseIndex(layoutMatch);
                return new PackageScanEntry(entry, PresentationScope.CustomLayout, null, "Custom layout " + index, 1000);
            }

            var notesMatch = NotesMasterPattern.Match(fullName);
            if (notesMatch.Success)
            {
                var index = ParseIndex(notesMatch);
                var label = index == 1 ? "Notes master" : "Notes master " + index;
                return new PackageScanEntry(entry, PresentationScope.NotesMaster, null, label, 3000);
            }

            return null;
        }

        private static int ParseIndex(Match match)
        {
            if (match == null || !match.Success)
            {
                return 0;
            }

            int index;
            return int.TryParse(match.Groups[1].Value, out index) ? index : 0;
        }

        private static bool TryParseOfficeBooleanAttribute(string rawValue, out bool value)
        {
            value = false;
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                return false;
            }

            if (bool.TryParse(rawValue, out value))
            {
                return true;
            }

            int numericValue;
            if (int.TryParse(rawValue, out numericValue))
            {
                value = numericValue != 0;
                return true;
            }

            return false;
        }

        private static void ScanFontUsages(
            Stream stream,
            PackageScanEntry scanEntry,
            ICollection<PackageFontUsageRecord> results)
        {
            if (stream == null || scanEntry == null || results == null)
            {
                return;
            }

            using (var reader = CreateReader(stream))
            {
                var frameStack = new Stack<XmlScanFrame>();
                var shapeStack = new Stack<FontShapeContext>();

                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        var isShapeBoundary = IsShapeElement(reader.NamespaceURI, reader.LocalName);
                        if (isShapeBoundary)
                        {
                            shapeStack.Push(new FontShapeContext());
                        }

                        frameStack.Push(new XmlScanFrame(reader.NamespaceURI, reader.LocalName, isShapeBoundary));

                        if (reader.NamespaceURI == PresentationNamespace.NamespaceName
                            && reader.LocalName.Equals("cNvPr", StringComparison.OrdinalIgnoreCase)
                            && shapeStack.Count > 0)
                        {
                            shapeStack.Peek().ApplyShapeProperties(reader.GetAttribute("name"), reader.GetAttribute("id"));
                        }
                        else if (shapeStack.Count > 0
                            && reader.NamespaceURI == DrawingNamespace.NamespaceName
                            && IsFontTypefaceElement(reader.LocalName))
                        {
                            var typeface = NormalizeFontName(reader.GetAttribute("typeface"));
                            if (!string.IsNullOrWhiteSpace(typeface))
                            {
                                shapeStack.Peek().Fonts.Add(typeface);
                            }
                        }

                        if (reader.IsEmptyElement)
                        {
                            PopFontFrame(frameStack, shapeStack, scanEntry, results);
                        }
                    }
                    else if (reader.NodeType == XmlNodeType.EndElement)
                    {
                        PopFontFrame(frameStack, shapeStack, scanEntry, results);
                    }
                }
            }
        }

        private static void PopEmbeddedFontFrame(Stack<XmlScanFrame> frameStack, ref int embeddedFontListDepth)
        {
            if (frameStack == null || frameStack.Count == 0)
            {
                return;
            }

            var poppedDepth = frameStack.Count;
            frameStack.Pop();

            if (embeddedFontListDepth == poppedDepth)
            {
                embeddedFontListDepth = 0;
            }
        }

        private static void PopThemeFontFrame(Stack<XmlScanFrame> frameStack, ref string currentScheme)
        {
            if (frameStack == null || frameStack.Count == 0)
            {
                return;
            }

            var frame = frameStack.Pop();
            if (frame.NamespaceUri == DrawingNamespace.NamespaceName
                && (frame.LocalName.Equals("majorFont", StringComparison.OrdinalIgnoreCase)
                    || frame.LocalName.Equals("minorFont", StringComparison.OrdinalIgnoreCase)))
            {
                currentScheme = string.Empty;
            }
        }

        private static void PopFontFrame(
            Stack<XmlScanFrame> frameStack,
            Stack<FontShapeContext> shapeStack,
            PackageScanEntry scanEntry,
            ICollection<PackageFontUsageRecord> results)
        {
            if (frameStack == null || frameStack.Count == 0)
            {
                return;
            }

            var frame = frameStack.Pop();
            if (!frame.IsShapeBoundary || shapeStack == null || shapeStack.Count == 0)
            {
                return;
            }

            var shapeContext = shapeStack.Pop();
            if (shapeContext.Fonts.Count == 0)
            {
                return;
            }

            var shapeName = !string.IsNullOrWhiteSpace(shapeContext.ShapeName)
                ? shapeContext.ShapeName
                : (shapeContext.ShapeId.HasValue ? "Shape " + shapeContext.ShapeId.Value : "Shape");

            foreach (var fontName in shapeContext.Fonts)
            {
                results.Add(new PackageFontUsageRecord
                {
                    FontName = fontName,
                    Location = BuildLocation(scanEntry, shapeName, shapeContext.ShapeId)
                });
            }
        }

        private static bool IsShapeElement(string namespaceUri, string localName)
        {
            if (!string.Equals(namespaceUri, PresentationNamespace.NamespaceName, StringComparison.Ordinal))
            {
                return false;
            }

            return localName.Equals("sp", StringComparison.OrdinalIgnoreCase)
                || localName.Equals("cxnSp", StringComparison.OrdinalIgnoreCase)
                || localName.Equals("graphicFrame", StringComparison.OrdinalIgnoreCase)
                || localName.Equals("pic", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsFontTypefaceElement(string localName)
        {
            return localName.Equals("latin", StringComparison.OrdinalIgnoreCase)
                || localName.Equals("ea", StringComparison.OrdinalIgnoreCase)
                || localName.Equals("cs", StringComparison.OrdinalIgnoreCase)
                || localName.Equals("sym", StringComparison.OrdinalIgnoreCase)
                || localName.Equals("buFont", StringComparison.OrdinalIgnoreCase);
        }

        private static XmlReader CreateReader(Stream stream)
        {
            return XmlReader.Create(stream, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                IgnoreWhitespace = true
            });
        }

        private sealed class XmlScanFrame
        {
            public XmlScanFrame(string namespaceUri, string localName, bool isShapeBoundary)
            {
                NamespaceUri = namespaceUri ?? string.Empty;
                LocalName = localName ?? string.Empty;
                IsShapeBoundary = isShapeBoundary;
            }

            public string NamespaceUri { get; }

            public string LocalName { get; }

            public bool IsShapeBoundary { get; }
        }

        private sealed class FontShapeContext
        {
            public FontShapeContext()
            {
                Fonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            public string ShapeName { get; private set; }

            public int? ShapeId { get; private set; }

            public ISet<string> Fonts { get; }

            public void ApplyShapeProperties(string name, string rawId)
            {
                if (string.IsNullOrWhiteSpace(ShapeName))
                {
                    ShapeName = string.IsNullOrWhiteSpace(name) ? null : name.Trim();
                }

                if (!ShapeId.HasValue)
                {
                    int parsedId;
                    if (int.TryParse(rawId, out parsedId))
                    {
                        ShapeId = parsedId;
                    }
                }
            }
        }

        private sealed class PackageScanEntry
        {
            public PackageScanEntry(ZipArchiveEntry entry, PresentationScope scope, int? index, string label, int sortBucket)
            {
                Entry = entry;
                Scope = scope;
                Index = index;
                Label = label;
                SortOrder = sortBucket + (index ?? 0);
            }

            public ZipArchiveEntry Entry { get; }

            public PresentationScope Scope { get; }

            public int? Index { get; }

            public string Label { get; }

            public int SortOrder { get; }
        }
    }
}
