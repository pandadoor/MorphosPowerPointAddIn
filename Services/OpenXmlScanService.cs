using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using MorphosPowerPointAddIn.Models;
using MorphosPowerPointAddIn.Utilities;

namespace MorphosPowerPointAddIn.Services
{
    internal sealed class OpenXmlPackageScanResult
    {
        public OpenXmlPackageScanResult(
            IReadOnlyList<PackageFontUsageRecord> fontUsages,
            ColorScanResult colorScanResult,
            FontTrie fontTrie)
        {
            FontUsages = fontUsages ?? Array.Empty<PackageFontUsageRecord>();
            ColorScanResult = colorScanResult ?? new ColorScanResult();
            FontTrie = fontTrie ?? new FontTrie();
        }

        public IReadOnlyList<PackageFontUsageRecord> FontUsages { get; }

        public ColorScanResult ColorScanResult { get; }

        public FontTrie FontTrie { get; }
    }

    internal sealed class OpenXmlScanService
    {
        private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string PresentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
        private const string ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        private readonly AhoCorasickMatcher<UsageMarker> _usageMatcher;
        private readonly AhoCorasickMatcher<string> _fontElementMatcher;

        public OpenXmlScanService()
        {
            _usageMatcher = new AhoCorasickMatcher<UsageMarker>();
            foreach (var pattern in new[] { "/ln/", "/uln/", "/lnref/" })
            {
                _usageMatcher.Add(pattern, UsageMarker.Line);
            }

            foreach (var pattern in new[] { "/rpr/", "/defrpr/", "/endpararpr/", "/txpr/", "/txbody/", "/fontref/", "/highlight/" })
            {
                _usageMatcher.Add(pattern, UsageMarker.Text);
            }

            foreach (var pattern in new[] { "/effectlst/", "/effectstyle/", "/outershdw/", "/innershdw/", "/glow/", "/effectref/" })
            {
                _usageMatcher.Add(pattern, UsageMarker.Effect);
            }

            _usageMatcher.Build();

            _fontElementMatcher = new AhoCorasickMatcher<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var elementName in new[] { "latin", "ea", "cs", "sym", "buFont" })
            {
                _fontElementMatcher.Add(elementName, elementName);
            }

            _fontElementMatcher.Build();
        }

        public OpenXmlPackageScanResult ScanPackage(
            string filePath,
            IProgress<ScanProgressInfo> fontProgress,
            IProgress<ScanProgressInfo> colorProgress,
            CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return new OpenXmlPackageScanResult(Array.Empty<PackageFontUsageRecord>(), new ColorScanResult(), new FontTrie());
            }

            using (var document = PresentationDocument.Open(filePath, false))
            {
                var descriptors = BuildScanPartDescriptors(document);
                var fontTrie = new FontTrie();
                var fontUsages = new List<PackageFontUsageRecord>();
                var colorNodes = new ConcurrentDictionary<int, SharedColorNode>();
                var themeColors = ReadThemeColors(document);

                for (var i = 0; i < descriptors.Count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var descriptor = descriptors[i];
                    var startProgress = new ScanProgressInfo
                    {
                        CompletedItems = i,
                        TotalItems = descriptors.Count,
                        Message = "Indexing " + descriptor.Label
                    };

                    fontProgress?.Report(startProgress);
                    colorProgress?.Report(startProgress);

                    using (var stream = descriptor.Part.GetStream(FileMode.Open, FileAccess.Read))
                    {
                        ScanPart(stream, descriptor, fontUsages, fontTrie, colorNodes);
                    }

                    var endProgress = new ScanProgressInfo
                    {
                        CompletedItems = i + 1,
                        TotalItems = descriptors.Count,
                        Message = "Indexed " + descriptor.Label
                    };

                    fontProgress?.Report(endProgress);
                    colorProgress?.Report(endProgress);
                }

                return new OpenXmlPackageScanResult(
                    fontUsages,
                    BuildColorScanResult(colorNodes, themeColors),
                    fontTrie);
            }
        }

        public IReadOnlyList<PackageFontUsageRecord> ReadFontUsages(
            string filePath,
            IProgress<ScanProgressInfo> progress,
            CancellationToken cancellationToken)
        {
            return ScanPackage(filePath, progress, null, cancellationToken).FontUsages;
        }

        public ColorScanResult ReadDirectColorUsages(
            string filePath,
            IProgress<ScanProgressInfo> progress,
            CancellationToken cancellationToken)
        {
            return ScanPackage(filePath, null, progress, cancellationToken).ColorScanResult;
        }

        private void ScanPart(
            Stream stream,
            OpenXmlPartDescriptor descriptor,
            ICollection<PackageFontUsageRecord> fontUsages,
            FontTrie fontTrie,
            ConcurrentDictionary<int, SharedColorNode> colorNodes)
        {
            if (stream == null || descriptor == null || fontUsages == null || fontTrie == null || colorNodes == null)
            {
                return;
            }

            using (var reader = CreateReader(stream))
            {
                var frameStack = new Stack<XmlScanFrame>();
                var shapeStack = new Stack<ShapeScanContext>();
                var chartDepth = 0;

                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        var isShapeBoundary = IsShapeElement(reader.NamespaceURI, reader.LocalName);
                        if (isShapeBoundary)
                        {
                            shapeStack.Push(new ShapeScanContext());
                        }

                        var isChartElement = string.Equals(reader.NamespaceURI, ChartNamespace, StringComparison.Ordinal);
                        frameStack.Push(new XmlScanFrame(reader.NamespaceURI, reader.LocalName, isShapeBoundary, isChartElement));

                        if (isChartElement)
                        {
                            chartDepth++;
                        }

                        if (shapeStack.Count > 0
                            && string.Equals(reader.NamespaceURI, PresentationNamespace, StringComparison.Ordinal)
                            && reader.LocalName.Equals("cNvPr", StringComparison.OrdinalIgnoreCase))
                        {
                            shapeStack.Peek().ApplyShapeProperties(reader.GetAttribute("name"), reader.GetAttribute("id"));
                        }
                        else if (shapeStack.Count > 0
                            && string.Equals(reader.NamespaceURI, DrawingNamespace, StringComparison.Ordinal)
                            && _fontElementMatcher.Matches(reader.LocalName))
                        {
                            var typeface = NormalizeFontName(reader.GetAttribute("typeface"));
                            if (!string.IsNullOrWhiteSpace(typeface))
                            {
                                shapeStack.Peek().Fonts.Add(typeface);
                                fontTrie.Add(typeface);
                            }
                        }
                        else if (shapeStack.Count > 0
                            && chartDepth == 0
                            && string.Equals(reader.NamespaceURI, DrawingNamespace, StringComparison.Ordinal)
                            && IsDirectColorElement(reader.LocalName))
                        {
                            var hexValue = NormalizeColorHex(reader.GetAttribute("val") ?? reader.GetAttribute("lastClr"));
                            if (!string.IsNullOrWhiteSpace(hexValue))
                            {
                                RegisterColorUsage(descriptor, frameStack, shapeStack.Peek(), colorNodes, hexValue);
                            }
                        }

                        if (reader.IsEmptyElement)
                        {
                            PopFrame(frameStack, shapeStack, descriptor, fontUsages, ref chartDepth);
                        }
                    }
                    else if (reader.NodeType == XmlNodeType.EndElement)
                    {
                        PopFrame(frameStack, shapeStack, descriptor, fontUsages, ref chartDepth);
                    }
                }
            }
        }

        private void RegisterColorUsage(
            OpenXmlPartDescriptor descriptor,
            Stack<XmlScanFrame> frameStack,
            ShapeScanContext shapeContext,
            ConcurrentDictionary<int, SharedColorNode> colorNodes,
            string hexValue)
        {
            if (descriptor == null || shapeContext == null || colorNodes == null || string.IsNullOrWhiteSpace(hexValue))
            {
                return;
            }

            var usageKind = ResolveUsageKind(frameStack);
            if (usageKind == ColorUsageKind.ChartOverride)
            {
                return;
            }

            var seenKey = ((int)usageKind).ToString() + "|" + hexValue;
            if (!shapeContext.SeenColorKeys.Add(seenKey))
            {
                return;
            }

            var argb = ToArgb(hexValue);
            var colorNode = colorNodes.GetOrAdd(argb, key => new SharedColorNode(hexValue));
            colorNode.RegisterUsage(usageKind, shapeContext.GetOrCreateLocation(descriptor));
        }

        private static void PopFrame(
            Stack<XmlScanFrame> frameStack,
            Stack<ShapeScanContext> shapeStack,
            OpenXmlPartDescriptor descriptor,
            ICollection<PackageFontUsageRecord> fontUsages,
            ref int chartDepth)
        {
            if (frameStack == null || frameStack.Count == 0)
            {
                return;
            }

            var frame = frameStack.Pop();
            if (frame.IsChartElement && chartDepth > 0)
            {
                chartDepth--;
            }

            if (!frame.IsShapeBoundary || shapeStack == null || shapeStack.Count == 0)
            {
                return;
            }

            var shapeContext = shapeStack.Pop();
            if (shapeContext.Fonts.Count == 0)
            {
                return;
            }

            foreach (var fontName in shapeContext.Fonts)
            {
                fontUsages.Add(new PackageFontUsageRecord
                {
                    FontName = fontName,
                    Location = shapeContext.GetOrCreateLocation(descriptor)
                });
            }
        }

        private ColorScanResult BuildColorScanResult(
            ConcurrentDictionary<int, SharedColorNode> colorNodes,
            IReadOnlyList<ThemeColorInfo> themeColors)
        {
            var themeLookup = (themeColors ?? Array.Empty<ThemeColorInfo>())
                .Where(x => x != null && !string.IsNullOrWhiteSpace(x.HexValue))
                .GroupBy(x => x.HexValue, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(x => x.Key, x => x.First(), StringComparer.OrdinalIgnoreCase);

            var items = new List<ColorInventoryItem>();
            foreach (var colorNode in colorNodes.Values)
            {
                foreach (var usage in colorNode.GetUsages())
                {
                    ThemeColorInfo themeColor;
                    var hasThemeMatch = themeLookup.TryGetValue(colorNode.HexValue, out themeColor);

                    items.Add(new ColorInventoryItem
                    {
                        UsageKind = usage.UsageKind,
                        UsageKindLabel = GetColorUsageLabel(usage.UsageKind),
                        HexValue = colorNode.HexValue,
                        RgbValue = BuildRgbValue(colorNode.HexValue),
                        UsesCount = usage.UsesCount,
                        MatchesThemeColor = hasThemeMatch,
                        MatchingThemeDisplayName = hasThemeMatch ? themeColor.DisplayName : string.Empty,
                        MatchingThemeSchemeName = hasThemeMatch ? themeColor.SchemeName : string.Empty,
                        Locations = usage.Locations
                    });
                }
            }

            return new ColorScanResult
            {
                Items = items
                    .OrderBy(x => x.UsageKindLabel, StringComparer.OrdinalIgnoreCase)
                    .ThenByDescending(x => x.UsesCount)
                    .ThenBy(x => x.HexValue, StringComparer.OrdinalIgnoreCase)
                    .ToList(),
                ThemeColors = themeColors ?? Array.Empty<ThemeColorInfo>()
            };
        }

        private IReadOnlyList<ThemeColorInfo> ReadThemeColors(PresentationDocument document)
        {
            if (document == null || document.PresentationPart == null)
            {
                return Array.Empty<ThemeColorInfo>();
            }

            var colors = new List<ThemeColorInfo>();
            var visitedUris = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var slideMasterPart in document.PresentationPart.SlideMasterParts)
            {
                if (slideMasterPart == null || slideMasterPart.ThemePart == null)
                {
                    continue;
                }

                var uri = slideMasterPart.ThemePart.Uri == null
                    ? string.Empty
                    : slideMasterPart.ThemePart.Uri.ToString();
                if (!visitedUris.Add(uri))
                {
                    continue;
                }

                using (var stream = slideMasterPart.ThemePart.GetStream(FileMode.Open, FileAccess.Read))
                {
                    ReadThemeColors(stream, colors);
                }
            }

            return colors
                .GroupBy(x => x.SchemeName ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .Select(x => x.First())
                .ToList();
        }

        private static void ReadThemeColors(Stream stream, ICollection<ThemeColorInfo> colors)
        {
            if (stream == null || colors == null)
            {
                return;
            }

            using (var reader = CreateReader(stream))
            {
                var frameStack = new Stack<string>();
                var colorSchemeDepth = 0;
                var schemeEntryDepth = 0;
                string schemeName = null;

                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        frameStack.Push(reader.LocalName);

                        if (string.Equals(reader.NamespaceURI, DrawingNamespace, StringComparison.Ordinal)
                            && reader.LocalName.Equals("clrScheme", StringComparison.OrdinalIgnoreCase))
                        {
                            colorSchemeDepth = frameStack.Count;
                        }
                        else if (colorSchemeDepth > 0
                            && frameStack.Count == colorSchemeDepth + 1
                            && string.Equals(reader.NamespaceURI, DrawingNamespace, StringComparison.Ordinal))
                        {
                            schemeName = reader.LocalName;
                            schemeEntryDepth = frameStack.Count;
                        }
                        else if (schemeEntryDepth > 0
                            && string.Equals(reader.NamespaceURI, DrawingNamespace, StringComparison.Ordinal)
                            && (reader.LocalName.Equals("srgbClr", StringComparison.OrdinalIgnoreCase)
                                || reader.LocalName.Equals("sysClr", StringComparison.OrdinalIgnoreCase)))
                        {
                            var hexValue = NormalizeColorHex(reader.GetAttribute("val") ?? reader.GetAttribute("lastClr"));
                            if (!string.IsNullOrWhiteSpace(hexValue)
                                && !colors.Any(x => string.Equals(x.SchemeName, schemeName, StringComparison.OrdinalIgnoreCase)))
                            {
                                colors.Add(new ThemeColorInfo
                                {
                                    SchemeName = schemeName,
                                    DisplayName = GetThemeDisplayName(schemeName),
                                    HexValue = hexValue
                                });
                            }
                        }

                        if (reader.IsEmptyElement)
                        {
                            PopThemeFrame(frameStack, ref colorSchemeDepth, ref schemeEntryDepth, ref schemeName);
                        }
                    }
                    else if (reader.NodeType == XmlNodeType.EndElement)
                    {
                        PopThemeFrame(frameStack, ref colorSchemeDepth, ref schemeEntryDepth, ref schemeName);
                    }
                }
            }
        }

        private static void PopThemeFrame(
            Stack<string> frameStack,
            ref int colorSchemeDepth,
            ref int schemeEntryDepth,
            ref string schemeName)
        {
            if (frameStack == null || frameStack.Count == 0)
            {
                return;
            }

            var poppedDepth = frameStack.Count;
            frameStack.Pop();

            if (schemeEntryDepth == poppedDepth)
            {
                schemeEntryDepth = 0;
                schemeName = null;
            }

            if (colorSchemeDepth == poppedDepth)
            {
                colorSchemeDepth = 0;
            }
        }

        private ColorUsageKind ResolveUsageKind(IEnumerable<XmlScanFrame> frames)
        {
            if (frames == null)
            {
                return ColorUsageKind.ShapeFill;
            }

            if (frames.Any(x => x.IsChartElement))
            {
                return ColorUsageKind.ChartOverride;
            }

            var contextSignature = "/" + string.Join("/", frames.Reverse().Select(x => (x.LocalName ?? string.Empty).ToLowerInvariant())) + "/";
            var markers = _usageMatcher.Find(contextSignature);
            if (markers.Contains(UsageMarker.Line))
            {
                return ColorUsageKind.Line;
            }

            if (markers.Contains(UsageMarker.Text))
            {
                return ColorUsageKind.TextFill;
            }

            if (markers.Contains(UsageMarker.Effect))
            {
                return ColorUsageKind.Effect;
            }

            return ColorUsageKind.ShapeFill;
        }

        private static IReadOnlyList<OpenXmlPartDescriptor> BuildScanPartDescriptors(PresentationDocument document)
        {
            var descriptors = new List<OpenXmlPartDescriptor>();
            if (document == null || document.PresentationPart == null)
            {
                return descriptors;
            }

            var presentationPart = document.PresentationPart;

            var slideMasters = presentationPart.SlideMasterParts
                .OrderBy(x => ParsePartIndex(x.Uri))
                .ThenBy(x => x.Uri == null ? string.Empty : x.Uri.ToString(), StringComparer.OrdinalIgnoreCase)
                .ToList();

            for (var i = 0; i < slideMasters.Count; i++)
            {
                var slideMasterPart = slideMasters[i];
                descriptors.Add(new OpenXmlPartDescriptor(
                    slideMasterPart,
                    PresentationScope.SlideMaster,
                    null,
                    i == 0 ? "Slide master" : "Slide master " + (i + 1)));

                var layouts = slideMasterPart.SlideLayoutParts
                    .OrderBy(x => ParsePartIndex(x.Uri))
                    .ThenBy(x => x.Uri == null ? string.Empty : x.Uri.ToString(), StringComparer.OrdinalIgnoreCase)
                    .ToList();

                for (var layoutIndex = 0; layoutIndex < layouts.Count; layoutIndex++)
                {
                    descriptors.Add(new OpenXmlPartDescriptor(
                        layouts[layoutIndex],
                        PresentationScope.CustomLayout,
                        null,
                        "Custom layout " + (layoutIndex + 1)));
                }
            }

            var slideIndex = 1;
            var slideIds = presentationPart.Presentation == null || presentationPart.Presentation.SlideIdList == null
                ? Enumerable.Empty<SlideId>()
                : presentationPart.Presentation.SlideIdList.Elements<SlideId>();

            foreach (var slideId in slideIds)
            {
                if (slideId == null || slideId.RelationshipId == null)
                {
                    continue;
                }

                var slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                if (slidePart == null)
                {
                    continue;
                }

                descriptors.Add(new OpenXmlPartDescriptor(
                    slidePart,
                    PresentationScope.Slide,
                    slideIndex,
                    "Slide " + slideIndex));
                slideIndex++;
            }

            if (presentationPart.NotesMasterPart != null)
            {
                descriptors.Add(new OpenXmlPartDescriptor(
                    presentationPart.NotesMasterPart,
                    PresentationScope.NotesMaster,
                    null,
                    "Notes master"));
            }

            return descriptors;
        }

        private static bool IsShapeElement(string namespaceUri, string localName)
        {
            if (!string.Equals(namespaceUri, PresentationNamespace, StringComparison.Ordinal))
            {
                return false;
            }

            return localName.Equals("sp", StringComparison.OrdinalIgnoreCase)
                || localName.Equals("cxnSp", StringComparison.OrdinalIgnoreCase)
                || localName.Equals("graphicFrame", StringComparison.OrdinalIgnoreCase)
                || localName.Equals("pic", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsDirectColorElement(string localName)
        {
            return localName.Equals("srgbClr", StringComparison.OrdinalIgnoreCase)
                || localName.Equals("sysClr", StringComparison.OrdinalIgnoreCase);
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

        private static string NormalizeFontName(string fontName)
        {
            return string.IsNullOrWhiteSpace(fontName) ? string.Empty : fontName.Trim();
        }

        private static string NormalizeColorHex(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var normalized = value.Trim().TrimStart('#').ToUpperInvariant();
            return normalized.Length == 6 ? normalized : string.Empty;
        }

        private static int ToArgb(string hexValue)
        {
            var normalized = NormalizeColorHex(hexValue);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return 0;
            }

            return unchecked((int)(0xFF000000 | Convert.ToInt32(normalized, 16)));
        }

        private static string BuildRgbValue(string hexValue)
        {
            var normalized = NormalizeColorHex(hexValue);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return "0, 0, 0";
            }

            return Convert.ToInt32(normalized.Substring(0, 2), 16) + ", "
                + Convert.ToInt32(normalized.Substring(2, 2), 16) + ", "
                + Convert.ToInt32(normalized.Substring(4, 2), 16);
        }

        private static int ParsePartIndex(Uri uri)
        {
            if (uri == null)
            {
                return 0;
            }

            var text = Path.GetFileNameWithoutExtension(uri.ToString()) ?? string.Empty;
            var digits = new string(text.Reverse().TakeWhile(char.IsDigit).Reverse().ToArray());
            int index;
            return int.TryParse(digits, out index) ? index : 0;
        }

        private static string GetColorUsageLabel(ColorUsageKind usageKind)
        {
            switch (usageKind)
            {
                case ColorUsageKind.TextFill:
                    return "Text fill";
                case ColorUsageKind.Line:
                    return "Line";
                case ColorUsageKind.Effect:
                    return "Effect";
                case ColorUsageKind.ChartOverride:
                    return "Chart override";
                default:
                    return "Shape fill";
            }
        }

        private static string GetThemeDisplayName(string schemeName)
        {
            switch ((schemeName ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "dk1":
                    return "Dark 1";
                case "lt1":
                    return "Light 1";
                case "dk2":
                    return "Dark 2";
                case "lt2":
                    return "Light 2";
                case "accent1":
                    return "Accent 1";
                case "accent2":
                    return "Accent 2";
                case "accent3":
                    return "Accent 3";
                case "accent4":
                    return "Accent 4";
                case "accent5":
                    return "Accent 5";
                case "accent6":
                    return "Accent 6";
                case "hlink":
                    return "Hyperlink";
                case "folhlink":
                    return "Followed link";
                default:
                    return string.IsNullOrWhiteSpace(schemeName) ? "Theme color" : schemeName;
            }
        }

        private enum UsageMarker
        {
            Line,
            Text,
            Effect
        }

        private sealed class OpenXmlPartDescriptor
        {
            public OpenXmlPartDescriptor(OpenXmlPart part, PresentationScope scope, int? slideIndex, string label)
            {
                Part = part;
                Scope = scope;
                SlideIndex = slideIndex;
                Label = label ?? string.Empty;
            }

            public OpenXmlPart Part { get; }

            public PresentationScope Scope { get; }

            public int? SlideIndex { get; }

            public string Label { get; }
        }

        private sealed class XmlScanFrame
        {
            public XmlScanFrame(string namespaceUri, string localName, bool isShapeBoundary, bool isChartElement)
            {
                NamespaceUri = namespaceUri ?? string.Empty;
                LocalName = localName ?? string.Empty;
                IsShapeBoundary = isShapeBoundary;
                IsChartElement = isChartElement;
            }

            public string NamespaceUri { get; }

            public string LocalName { get; }

            public bool IsShapeBoundary { get; }

            public bool IsChartElement { get; }
        }

        private sealed class ShapeScanContext
        {
            private FontUsageLocation _location;

            public ShapeScanContext()
            {
                Fonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                SeenColorKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            public ISet<string> Fonts { get; }

            public ISet<string> SeenColorKeys { get; }

            public int? ShapeId { get; private set; }

            public string ShapeName { get; private set; }

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

            public FontUsageLocation GetOrCreateLocation(OpenXmlPartDescriptor descriptor)
            {
                if (_location != null)
                {
                    return _location;
                }

                var shapeName = !string.IsNullOrWhiteSpace(ShapeName)
                    ? ShapeName
                    : (ShapeId.HasValue ? "Shape " + ShapeId.Value : "Shape");

                var label = descriptor.Scope == PresentationScope.Slide && descriptor.SlideIndex.HasValue
                    ? "Slide " + descriptor.SlideIndex.Value + " - " + shapeName
                    : descriptor.Label + " - " + shapeName;

                _location = new FontUsageLocation
                {
                    Scope = descriptor.Scope,
                    SlideIndex = descriptor.Scope == PresentationScope.Slide ? descriptor.SlideIndex : null,
                    ShapeId = ShapeId,
                    ScopeName = descriptor.Label,
                    ShapeName = shapeName,
                    Label = label,
                    IsSelectable = descriptor.Scope == PresentationScope.Slide && descriptor.SlideIndex.HasValue
                };

                return _location;
            }
        }

        private sealed class SharedColorNode
        {
            private readonly ConcurrentDictionary<ColorUsageKind, ColorUsageAccumulator> _usages =
                new ConcurrentDictionary<ColorUsageKind, ColorUsageAccumulator>();

            public SharedColorNode(string hexValue)
            {
                HexValue = hexValue;
            }

            public string HexValue { get; }

            public void RegisterUsage(ColorUsageKind usageKind, FontUsageLocation location)
            {
                var usage = _usages.GetOrAdd(usageKind, kind => new ColorUsageAccumulator(kind));
                usage.Register(location);
            }

            public IReadOnlyList<ColorUsageAccumulatorSnapshot> GetUsages()
            {
                return _usages.Values
                    .Select(x => x.ToSnapshot())
                    .OrderBy(x => x.UsageKind)
                    .ToList();
            }
        }

        private sealed class ColorUsageAccumulator
        {
            private readonly object _sync = new object();
            private readonly Dictionary<string, FontUsageLocation> _locations =
                new Dictionary<string, FontUsageLocation>(StringComparer.OrdinalIgnoreCase);

            public ColorUsageAccumulator(ColorUsageKind usageKind)
            {
                UsageKind = usageKind;
            }

            public ColorUsageKind UsageKind { get; }

            public void Register(FontUsageLocation location)
            {
                if (location == null)
                {
                    return;
                }

                var key = BuildLocationKey(location);
                lock (_sync)
                {
                    if (_locations.ContainsKey(key))
                    {
                        return;
                    }

                    _locations[key] = location;
                }
            }

            public ColorUsageAccumulatorSnapshot ToSnapshot()
            {
                lock (_sync)
                {
                    return new ColorUsageAccumulatorSnapshot(
                        UsageKind,
                        _locations.Count,
                        _locations.Values.ToList());
                }
            }

            private static string BuildLocationKey(FontUsageLocation location)
            {
                return (location.Scope.ToString() ?? string.Empty) + "|"
                    + (location.SlideIndex.HasValue ? location.SlideIndex.Value.ToString() : string.Empty) + "|"
                    + (location.ShapeId.HasValue ? location.ShapeId.Value.ToString() : string.Empty) + "|"
                    + (location.ShapeName ?? string.Empty);
            }
        }

        private sealed class ColorUsageAccumulatorSnapshot
        {
            public ColorUsageAccumulatorSnapshot(
                ColorUsageKind usageKind,
                int usesCount,
                IReadOnlyList<FontUsageLocation> locations)
            {
                UsageKind = usageKind;
                UsesCount = usesCount;
                Locations = locations ?? Array.Empty<FontUsageLocation>();
            }

            public ColorUsageKind UsageKind { get; }

            public int UsesCount { get; }

            public IReadOnlyList<FontUsageLocation> Locations { get; }
        }
    }
}
