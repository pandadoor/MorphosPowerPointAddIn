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
    internal static class PresentationColorInspector
    {
        private const string ThemeEntryPath = "ppt/theme/theme1.xml";
        private static readonly Regex SlidePattern = new Regex(@"^ppt/slides/slide(\d+)\.xml$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly XNamespace PresentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
        private static readonly XNamespace DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static readonly XNamespace ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        public static ColorScanResult ReadDirectColorUsages(string filePath, IProgress<ScanProgressInfo> progress, CancellationToken cancellationToken)
        {
            try
            {
                return new Services.OpenXmlScanService()
                    .ReadDirectColorUsages(filePath, progress, cancellationToken);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch
            {
                return new ColorScanResult();
            }
        }

        public static int ApplyColorReplacements(string filePath, IReadOnlyList<ColorReplacementInstruction> instructions)
        {
            return new Services.OpenXmlColorReplacer()
                .ApplyColorReplacements(filePath, instructions);
        }

        private static IReadOnlyList<ThemeColorInfo> ReadThemeColors(ZipArchive archive)
        {
            var colors = new List<ThemeColorInfo>();
            if (archive == null)
            {
                return colors;
            }

            var entry = archive.Entries.FirstOrDefault(x => x.FullName.Equals(ThemeEntryPath, StringComparison.OrdinalIgnoreCase));
            if (entry == null)
            {
                return colors;
            }

            try
            {
                using (var stream = entry.Open())
                {
                    using (var reader = CreateReader(stream))
                    {
                        var frameStack = new Stack<XmlScanFrame>();
                        var colorSchemeDepth = 0;
                        var schemeEntryDepth = 0;
                        string schemeName = null;

                        while (reader.Read())
                        {
                            if (reader.NodeType == XmlNodeType.Element)
                            {
                                frameStack.Push(new XmlScanFrame(reader.NamespaceURI, reader.LocalName, false, false));

                                if (reader.NamespaceURI == DrawingNamespace.NamespaceName
                                    && reader.LocalName.Equals("clrScheme", StringComparison.OrdinalIgnoreCase))
                                {
                                    colorSchemeDepth = frameStack.Count;
                                }
                                else if (colorSchemeDepth > 0
                                    && frameStack.Count == colorSchemeDepth + 1
                                    && reader.NamespaceURI == DrawingNamespace.NamespaceName)
                                {
                                    schemeName = reader.LocalName;
                                    schemeEntryDepth = frameStack.Count;
                                }
                                else if (schemeEntryDepth > 0
                                    && reader.NamespaceURI == DrawingNamespace.NamespaceName
                                    && (reader.LocalName.Equals("srgbClr", StringComparison.OrdinalIgnoreCase)
                                        || reader.LocalName.Equals("sysClr", StringComparison.OrdinalIgnoreCase)))
                                {
                                    var hexValue = NormalizeHex(
                                        reader.GetAttribute("val")
                                        ?? reader.GetAttribute("lastClr"));
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
            }
            catch
            {
                return colors;
            }

            return colors;
        }

        private static void ScanSlideColors(Stream stream, int slideIndex, IDictionary<string, ColorAccumulator> accumulators)
        {
            if (stream == null || accumulators == null)
            {
                return;
            }

            using (var reader = CreateReader(stream))
            {
                var frameStack = new Stack<XmlScanFrame>();
                var shapeStack = new Stack<ColorShapeContext>();
                var chartDepth = 0;

                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        var isShapeBoundary = IsShapeElement(reader.NamespaceURI, reader.LocalName);
                        if (isShapeBoundary)
                        {
                            shapeStack.Push(new ColorShapeContext());
                        }

                        frameStack.Push(new XmlScanFrame(
                            reader.NamespaceURI,
                            reader.LocalName,
                            isShapeBoundary,
                            reader.NamespaceURI == ChartNamespace.NamespaceName));

                        if (reader.NamespaceURI == ChartNamespace.NamespaceName)
                        {
                            chartDepth++;
                        }

                        if (reader.NamespaceURI == PresentationNamespace.NamespaceName
                            && reader.LocalName.Equals("cNvPr", StringComparison.OrdinalIgnoreCase)
                            && shapeStack.Count > 0)
                        {
                            shapeStack.Peek().ApplyShapeProperties(reader.GetAttribute("name"), reader.GetAttribute("id"));
                        }
                        else if (reader.NamespaceURI == DrawingNamespace.NamespaceName
                            && reader.LocalName.Equals("srgbClr", StringComparison.OrdinalIgnoreCase)
                            && chartDepth == 0
                            && shapeStack.Count > 0)
                        {
                            var hexValue = NormalizeHex(reader.GetAttribute("val"));
                            if (!string.IsNullOrWhiteSpace(hexValue))
                            {
                                var shapeContext = shapeStack.Peek();
                                AddUsage(
                                    accumulators,
                                    ResolveUsageKind(frameStack),
                                    hexValue,
                                    shapeContext.GetOrCreateLocation(slideIndex),
                                    shapeContext.SeenKeys);
                            }
                        }

                        if (reader.IsEmptyElement)
                        {
                            PopColorFrame(frameStack, shapeStack, ref chartDepth);
                        }
                    }
                    else if (reader.NodeType == XmlNodeType.EndElement)
                    {
                        PopColorFrame(frameStack, shapeStack, ref chartDepth);
                    }
                }
            }
        }

        private static void PopThemeFrame(
            Stack<XmlScanFrame> frameStack,
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

        private static void PopColorFrame(
            Stack<XmlScanFrame> frameStack,
            Stack<ColorShapeContext> shapeStack,
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

            if (frame.IsShapeBoundary && shapeStack != null && shapeStack.Count > 0)
            {
                shapeStack.Pop();
            }
        }

        private static ColorUsageKind ResolveUsageKind(IEnumerable<XmlScanFrame> frames)
        {
            if (frames == null)
            {
                return ColorUsageKind.ShapeFill;
            }

            if (frames.Any(x => x.IsChartElement))
            {
                return ColorUsageKind.ChartOverride;
            }

            if (frames.Any(x =>
                x.LocalName.Equals("ln", StringComparison.OrdinalIgnoreCase)
                || x.LocalName.Equals("uLn", StringComparison.OrdinalIgnoreCase)
                || x.LocalName.Equals("lnRef", StringComparison.OrdinalIgnoreCase)))
            {
                return ColorUsageKind.Line;
            }

            if (frames.Any(x =>
                x.LocalName.Equals("rPr", StringComparison.OrdinalIgnoreCase)
                || x.LocalName.Equals("defRPr", StringComparison.OrdinalIgnoreCase)
                || x.LocalName.Equals("endParaRPr", StringComparison.OrdinalIgnoreCase)
                || x.LocalName.Equals("txPr", StringComparison.OrdinalIgnoreCase)
                || x.LocalName.Equals("txBody", StringComparison.OrdinalIgnoreCase)
                || x.LocalName.Equals("fontRef", StringComparison.OrdinalIgnoreCase)
                || x.LocalName.Equals("highlight", StringComparison.OrdinalIgnoreCase)))
            {
                return ColorUsageKind.TextFill;
            }

            if (frames.Any(x =>
                x.LocalName.Equals("effectLst", StringComparison.OrdinalIgnoreCase)
                || x.LocalName.Equals("effectStyle", StringComparison.OrdinalIgnoreCase)
                || x.LocalName.Equals("outerShdw", StringComparison.OrdinalIgnoreCase)
                || x.LocalName.Equals("innerShdw", StringComparison.OrdinalIgnoreCase)
                || x.LocalName.Equals("glow", StringComparison.OrdinalIgnoreCase)
                || x.LocalName.Equals("effectRef", StringComparison.OrdinalIgnoreCase)))
            {
                return ColorUsageKind.Effect;
            }

            return ColorUsageKind.ShapeFill;
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

        private static void ScanShapeContainer(IEnumerable<XElement> elements, int slideIndex, IDictionary<string, ColorAccumulator> accumulators)
        {
            foreach (var element in elements)
            {
                if (element == null || element.Name.Namespace != PresentationNamespace)
                {
                    continue;
                }

                if (element.Name == PresentationNamespace + "grpSp")
                {
                    ScanShapeContainer(element.Elements(), slideIndex, accumulators);
                    continue;
                }

                if (!IsShapeElement(element))
                {
                    continue;
                }

                var shapeName = ResolveShapeName(element);
                var shapeId = ResolveShapeId(element);
                var location = new FontUsageLocation
                {
                    Scope = PresentationScope.Slide,
                    SlideIndex = slideIndex,
                    ShapeId = shapeId,
                    ShapeName = shapeName,
                    ScopeName = "Slide " + slideIndex,
                    Label = "Slide " + slideIndex + " - " + shapeName,
                    IsSelectable = true
                };

                var seenInShape = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var colorElement in element.Descendants(DrawingNamespace + "srgbClr"))
                {
                    if (colorElement.Ancestors().Any(x => x.Name.Namespace == ChartNamespace))
                    {
                        continue;
                    }

                    var hexValue = NormalizeHex((string)colorElement.Attribute("val"));
                    if (string.IsNullOrWhiteSpace(hexValue))
                    {
                        continue;
                    }

                    var usageKind = ResolveUsageKind(colorElement);
                    if (usageKind == ColorUsageKind.ChartOverride)
                    {
                        continue;
                    }

                    AddUsage(accumulators, usageKind, hexValue, location, seenInShape);
                }
            }
        }

        private static bool ApplyColorReplacements(
            IEnumerable<XElement> elements,
            IDictionary<string, ColorReplacementInstruction> lookup,
            IDictionary<string, ColorReplacementInstruction> hexFallbackLookup,
            ref int replacements)
        {
            var changed = false;
            foreach (var element in elements)
            {
                if (element == null || element.Name.Namespace != PresentationNamespace)
                {
                    continue;
                }

                if (element.Name == PresentationNamespace + "grpSp")
                {
                    if (ApplyColorReplacements(element.Elements(), lookup, hexFallbackLookup, ref replacements))
                    {
                        changed = true;
                    }

                    continue;
                }

                if (!IsShapeElement(element))
                {
                    continue;
                }

                foreach (var colorElement in element.Descendants(DrawingNamespace + "srgbClr").ToList())
                {
                    if (colorElement.Ancestors().Any(x => x.Name.Namespace == ChartNamespace))
                    {
                        continue;
                    }

                    var hexValue = NormalizeHex((string)colorElement.Attribute("val"));
                    if (string.IsNullOrWhiteSpace(hexValue))
                    {
                        continue;
                    }

                    ColorReplacementInstruction instruction;
                    if (!lookup.TryGetValue(BuildInstructionKey(ResolveUsageKind(colorElement), hexValue), out instruction)
                        && !hexFallbackLookup.TryGetValue(NormalizeHex(hexValue), out instruction))
                    {
                        continue;
                    }

                    ReplaceColorElement(colorElement, instruction);
                    changed = true;
                    replacements++;
                }
            }

            return changed;
        }

        private static void ReplaceColorElement(XElement colorElement, ColorReplacementInstruction instruction)
        {
            if (colorElement == null || instruction == null)
            {
                return;
            }

            if (instruction.UseThemeColor && !string.IsNullOrWhiteSpace(instruction.ThemeSchemeName))
            {
                colorElement.ReplaceWith(new XElement(DrawingNamespace + "schemeClr", new XAttribute("val", instruction.ThemeSchemeName)));
                return;
            }

            if (!string.IsNullOrWhiteSpace(instruction.ReplacementHexValue))
            {
                colorElement.SetAttributeValue("val", NormalizeHex(instruction.ReplacementHexValue));
                colorElement.RemoveNodes();
            }
        }

        private static void AddUsage(
            IDictionary<string, ColorAccumulator> accumulators,
            ColorUsageKind usageKind,
            string hexValue,
            FontUsageLocation location,
            ISet<string> seenInShape)
        {
            var key = BuildInstructionKey(usageKind, hexValue);
            if (seenInShape != null && !seenInShape.Add(key))
            {
                return;
            }

            ColorAccumulator accumulator;
            if (!accumulators.TryGetValue(key, out accumulator))
            {
                accumulator = new ColorAccumulator(usageKind, hexValue);
                accumulators[key] = accumulator;
            }

            accumulator.UsesCount++;
            accumulator.Locations.Add(location);
        }

        private static bool IsValidInstruction(ColorReplacementInstruction instruction)
        {
            if (instruction == null || string.IsNullOrWhiteSpace(instruction.SourceHexValue))
            {
                return false;
            }

            if (instruction.UseThemeColor)
            {
                return !string.IsNullOrWhiteSpace(instruction.ThemeSchemeName);
            }

            return !string.IsNullOrWhiteSpace(instruction.ReplacementHexValue);
        }

        private static string BuildReplacementFingerprint(ColorReplacementInstruction instruction)
        {
            if (instruction == null)
            {
                return string.Empty;
            }

            if (instruction.UseThemeColor)
            {
                return "theme|" + (instruction.ThemeSchemeName ?? string.Empty).Trim();
            }

            return "rgb|" + NormalizeHex(instruction.ReplacementHexValue);
        }

        private static bool IsShapeElement(XElement element)
        {
            return element.Name == PresentationNamespace + "sp"
                || element.Name == PresentationNamespace + "cxnSp"
                || element.Name == PresentationNamespace + "graphicFrame"
                || element.Name == PresentationNamespace + "pic";
        }

        private static ColorUsageKind ResolveUsageKind(XElement colorElement)
        {
            if (colorElement == null)
            {
                return ColorUsageKind.ShapeFill;
            }

            if (colorElement.Ancestors().Any(x => x.Name.Namespace == ChartNamespace))
            {
                return ColorUsageKind.ChartOverride;
            }

            if (colorElement.Ancestors().Any(x =>
                x.Name.LocalName.Equals("ln", StringComparison.OrdinalIgnoreCase)
                || x.Name.LocalName.Equals("uLn", StringComparison.OrdinalIgnoreCase)
                || x.Name.LocalName.Equals("lnRef", StringComparison.OrdinalIgnoreCase)))
            {
                return ColorUsageKind.Line;
            }

            if (colorElement.Ancestors().Any(x =>
                x.Name.LocalName.Equals("rPr", StringComparison.OrdinalIgnoreCase)
                || x.Name.LocalName.Equals("defRPr", StringComparison.OrdinalIgnoreCase)
                || x.Name.LocalName.Equals("endParaRPr", StringComparison.OrdinalIgnoreCase)
                || x.Name.LocalName.Equals("txPr", StringComparison.OrdinalIgnoreCase)
                || x.Name.LocalName.Equals("txBody", StringComparison.OrdinalIgnoreCase)
                || x.Name.LocalName.Equals("fontRef", StringComparison.OrdinalIgnoreCase)
                || x.Name.LocalName.Equals("highlight", StringComparison.OrdinalIgnoreCase)))
            {
                return ColorUsageKind.TextFill;
            }

            if (colorElement.Ancestors().Any(x =>
                x.Name.LocalName.Equals("effectLst", StringComparison.OrdinalIgnoreCase)
                || x.Name.LocalName.Equals("effectStyle", StringComparison.OrdinalIgnoreCase)
                || x.Name.LocalName.Equals("outerShdw", StringComparison.OrdinalIgnoreCase)
                || x.Name.LocalName.Equals("innerShdw", StringComparison.OrdinalIgnoreCase)
                || x.Name.LocalName.Equals("glow", StringComparison.OrdinalIgnoreCase)
                || x.Name.LocalName.Equals("effectRef", StringComparison.OrdinalIgnoreCase)))
            {
                return ColorUsageKind.Effect;
            }

            return ColorUsageKind.ShapeFill;
        }

        private static string BuildInstructionKey(ColorUsageKind usageKind, string hexValue)
        {
            return ((int)usageKind).ToString() + "|" + NormalizeHex(hexValue);
        }

        private static string NormalizeHex(string hexValue)
        {
            if (string.IsNullOrWhiteSpace(hexValue))
            {
                return string.Empty;
            }

            var normalized = hexValue.Trim().TrimStart('#').ToUpperInvariant();
            return normalized.Length == 6 ? normalized : string.Empty;
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

            int parsedId;
            return int.TryParse((string)properties.Attribute("id"), out parsedId) ? parsedId : (int?)null;
        }

        private static int ParseSlideIndex(Match match)
        {
            if (match == null || !match.Success)
            {
                return 0;
            }

            int slideIndex;
            return int.TryParse(match.Groups[1].Value, out slideIndex) ? slideIndex : 0;
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

        private sealed class ColorShapeContext
        {
            private FontUsageLocation _location;

            public ColorShapeContext()
            {
                SeenKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            public int? ShapeId { get; private set; }

            public string ShapeName { get; private set; }

            public ISet<string> SeenKeys { get; }

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

            public FontUsageLocation GetOrCreateLocation(int slideIndex)
            {
                if (_location != null)
                {
                    return _location;
                }

                var resolvedShapeName = !string.IsNullOrWhiteSpace(ShapeName)
                    ? ShapeName
                    : (ShapeId.HasValue ? "Shape " + ShapeId.Value : "Shape");

                _location = new FontUsageLocation
                {
                    Scope = PresentationScope.Slide,
                    SlideIndex = slideIndex,
                    ShapeId = ShapeId,
                    ShapeName = resolvedShapeName,
                    ScopeName = "Slide " + slideIndex,
                    Label = "Slide " + slideIndex + " - " + resolvedShapeName,
                    IsSelectable = true
                };

                return _location;
            }
        }

        private sealed class ColorAccumulator
        {
            public ColorAccumulator(ColorUsageKind usageKind, string hexValue)
            {
                UsageKind = usageKind;
                HexValue = hexValue;
                Locations = new List<FontUsageLocation>();
            }

            public ColorUsageKind UsageKind { get; }

            public string HexValue { get; }

            public int UsesCount { get; set; }

            public List<FontUsageLocation> Locations { get; }

            public string UsageKindLabel
            {
                get
                {
                    switch (UsageKind)
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
            }

            public ColorInventoryItem ToInventoryItem(IDictionary<string, ThemeColorInfo> themeColorLookup)
            {
                ThemeColorInfo matchingTheme = null;
                var hasThemeMatch = themeColorLookup != null && themeColorLookup.TryGetValue(HexValue, out matchingTheme);

                return new ColorInventoryItem
                {
                    UsageKind = UsageKind,
                    UsageKindLabel = UsageKindLabel,
                    HexValue = HexValue,
                    RgbValue = HexToRgb(HexValue),
                    UsesCount = UsesCount,
                    MatchesThemeColor = hasThemeMatch,
                    MatchingThemeDisplayName = hasThemeMatch ? matchingTheme.DisplayName : string.Empty,
                    MatchingThemeSchemeName = hasThemeMatch ? matchingTheme.SchemeName : string.Empty,
                    Locations = Locations
                        .GroupBy(x => x.Label, StringComparer.OrdinalIgnoreCase)
                        .Select(x => x.First())
                        .ToList()
                };
            }

            private static string HexToRgb(string hexValue)
            {
                if (string.IsNullOrWhiteSpace(hexValue) || hexValue.Length != 6)
                {
                    return string.Empty;
                }

                var red = Convert.ToInt32(hexValue.Substring(0, 2), 16);
                var green = Convert.ToInt32(hexValue.Substring(2, 2), 16);
                var blue = Convert.ToInt32(hexValue.Substring(4, 2), 16);
                return red + ", " + green + ", " + blue;
            }
        }
    }
}
