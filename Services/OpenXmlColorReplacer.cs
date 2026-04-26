using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using MorphosPowerPointAddIn.Models;
using MorphosPowerPointAddIn.Utilities;

namespace MorphosPowerPointAddIn.Services
{
    internal sealed class OpenXmlColorReplacer
    {
        private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string PresentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
        private const string ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        public int ApplyColorReplacements(string filePath, IReadOnlyList<ColorReplacementInstruction> instructions)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath) || instructions == null || instructions.Count == 0)
            {
                return 0;
            }

            var sanitizedInstructions = instructions
                .Where(IsValidInstruction)
                .ToList();
            if (sanitizedInstructions.Count == 0)
            {
                return 0;
            }

            var lookup = sanitizedInstructions
                .GroupBy(x => BuildInstructionKey(x.UsageKind, x.SourceHexValue), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(x => x.Key, x => x.Last(), StringComparer.OrdinalIgnoreCase);

            var hexFallbackLookup = sanitizedInstructions
                .GroupBy(x => NormalizeHex(x.SourceHexValue), StringComparer.OrdinalIgnoreCase)
                .Where(x => x.Count() == 1 || x.Select(BuildReplacementFingerprint).Distinct(StringComparer.OrdinalIgnoreCase).Count() == 1)
                .ToDictionary(x => x.Key, x => x.Last(), StringComparer.OrdinalIgnoreCase);

            var matcher = BuildMatcher(sanitizedInstructions.Select(x => NormalizeHex(x.SourceHexValue)));
            var workItems = ReadCandidateParts(filePath, BuildSlidePartDescriptors, matcher);
            var results = ProcessQueue(
                workItems,
                item => ApplyColorReplacements(item, lookup, hexFallbackLookup));

            WriteResults(filePath, BuildSlidePartLookup, results);
            return results.Sum(x => x.ChangeCount);
        }

        public int ApplyFontReplacements(string filePath, IReadOnlyCollection<string> sourceFonts, string replacementFont)
        {
            if (string.IsNullOrWhiteSpace(filePath)
                || !File.Exists(filePath)
                || sourceFonts == null
                || sourceFonts.Count == 0
                || string.IsNullOrWhiteSpace(replacementFont))
            {
                return 0;
            }

            var normalizedFonts = new HashSet<string>(
                sourceFonts
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Select(x => x.Trim()),
                StringComparer.OrdinalIgnoreCase);
            if (normalizedFonts.Count == 0)
            {
                return 0;
            }

            var matcher = BuildMatcher(normalizedFonts.SelectMany(x => new[] { x, SecurityElement.Escape(x) ?? x }));
            var workItems = ReadCandidateParts(filePath, BuildFontPartDescriptors, matcher);
            var results = ProcessQueue(
                workItems,
                item => ApplyFontReplacements(item, normalizedFonts, replacementFont.Trim()));

            WriteResults(filePath, BuildFontPartLookup, results);
            return results.Sum(x => x.ChangeCount);
        }

        private static IReadOnlyList<XmlMutationResult> ProcessQueue(
            IReadOnlyList<XmlWorkItem> workItems,
            Func<XmlWorkItem, XmlMutationResult> mutator)
        {
            if (workItems == null || workItems.Count == 0 || mutator == null)
            {
                return Array.Empty<XmlMutationResult>();
            }

            using (var queue = new BlockingCollection<XmlWorkItem>(32))
            {
                var results = new ConcurrentBag<XmlMutationResult>();
                var consumer = Task.Run(() =>
                {
                    foreach (var workItem in queue.GetConsumingEnumerable())
                    {
                        var result = mutator(workItem);
                        if (result != null && result.ChangeCount > 0)
                        {
                            results.Add(result);
                        }
                    }
                });

                foreach (var workItem in workItems)
                {
                    queue.Add(workItem);
                }

                queue.CompleteAdding();
                consumer.Wait();

                return results
                    .OrderBy(x => x.SortOrder)
                    .ThenBy(x => x.PartUri, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
        }

        private static IReadOnlyList<XmlWorkItem> ReadCandidateParts(
            string filePath,
            Func<PresentationDocument, IReadOnlyList<PartDescriptor>> descriptorFactory,
            AhoCorasickMatcher<string> matcher)
        {
            var items = new List<XmlWorkItem>();
            using (var document = PresentationDocument.Open(filePath, false))
            {
                var descriptors = descriptorFactory(document);
                foreach (var descriptor in descriptors)
                {
                    using (var stream = descriptor.Part.GetStream(FileMode.Open, FileAccess.Read))
                    using (var reader = new StreamReader(stream))
                    {
                        var xml = reader.ReadToEnd();
                        if (string.IsNullOrWhiteSpace(xml))
                        {
                            continue;
                        }

                        if (matcher != null && !matcher.Matches(xml))
                        {
                            continue;
                        }

                        items.Add(new XmlWorkItem(descriptor.Uri, descriptor.SortOrder, xml));
                    }
                }
            }

            return items;
        }

        private static void WriteResults(
            string filePath,
            Func<PresentationDocument, IDictionary<string, OpenXmlPart>> lookupFactory,
            IReadOnlyList<XmlMutationResult> results)
        {
            if (string.IsNullOrWhiteSpace(filePath)
                || !File.Exists(filePath)
                || results == null
                || results.Count == 0
                || lookupFactory == null)
            {
                return;
            }

            using (var document = PresentationDocument.Open(filePath, true))
            {
                var lookup = lookupFactory(document);
                foreach (var result in results)
                {
                    OpenXmlPart part;
                    if (!lookup.TryGetValue(result.PartUri, out part))
                    {
                        continue;
                    }

                    using (var stream = part.GetStream(FileMode.Create, FileAccess.Write))
                    using (var writer = new StreamWriter(stream))
                    {
                        writer.Write(result.XmlText);
                    }
                }
            }
        }

        private static XmlMutationResult ApplyColorReplacements(
            XmlWorkItem workItem,
            IDictionary<string, ColorReplacementInstruction> lookup,
            IDictionary<string, ColorReplacementInstruction> hexFallbackLookup)
        {
            if (workItem == null || string.IsNullOrWhiteSpace(workItem.XmlText))
            {
                return null;
            }

            var document = XDocument.Parse(workItem.XmlText, LoadOptions.PreserveWhitespace);
            var replacements = 0;
            foreach (var colorElement in document
                .Descendants()
                .Where(IsDirectColorElement)
                .ToList())
            {
                if (colorElement.Ancestors().Any(x => x.Name.NamespaceName == ChartNamespace))
                {
                    continue;
                }

                var hexValue = NormalizeHex(
                    (string)colorElement.Attribute("val")
                    ?? (string)colorElement.Attribute("lastClr"));
                if (string.IsNullOrWhiteSpace(hexValue))
                {
                    continue;
                }

                ColorReplacementInstruction instruction;
                if (!lookup.TryGetValue(BuildInstructionKey(ResolveUsageKind(colorElement), hexValue), out instruction)
                    && !hexFallbackLookup.TryGetValue(hexValue, out instruction))
                {
                    continue;
                }

                ReplaceColorElement(colorElement, instruction);
                replacements++;
            }

            return replacements <= 0
                ? null
                : new XmlMutationResult(workItem.PartUri, workItem.SortOrder, document.ToString(SaveOptions.DisableFormatting), replacements);
        }

        private static XmlMutationResult ApplyFontReplacements(
            XmlWorkItem workItem,
            ISet<string> sourceFonts,
            string replacementFont)
        {
            if (workItem == null || string.IsNullOrWhiteSpace(workItem.XmlText) || sourceFonts == null || sourceFonts.Count == 0)
            {
                return null;
            }

            var document = XDocument.Parse(workItem.XmlText, LoadOptions.PreserveWhitespace);
            var replacements = 0;
            foreach (var fontElement in document
                .Descendants()
                .Where(IsFontElement)
                .ToList())
            {
                var typeface = NormalizeFontName((string)fontElement.Attribute("typeface"));
                if (!sourceFonts.Contains(typeface))
                {
                    continue;
                }

                fontElement.SetAttributeValue("typeface", replacementFont);
                replacements++;
            }

            return replacements <= 0
                ? null
                : new XmlMutationResult(workItem.PartUri, workItem.SortOrder, document.ToString(SaveOptions.DisableFormatting), replacements);
        }

        private static void ReplaceColorElement(XElement colorElement, ColorReplacementInstruction instruction)
        {
            if (colorElement == null || instruction == null)
            {
                return;
            }

            XName elementName;
            object[] content;
            if (instruction.UseThemeColor && !string.IsNullOrWhiteSpace(instruction.ThemeSchemeName))
            {
                elementName = XName.Get("schemeClr", DrawingNamespace);
                content = new object[] { new XAttribute("val", instruction.ThemeSchemeName.Trim()) };
            }
            else
            {
                var replacementHex = NormalizeHex(instruction.ReplacementHexValue);
                if (string.IsNullOrWhiteSpace(replacementHex))
                {
                    return;
                }

                elementName = XName.Get("srgbClr", DrawingNamespace);
                content = new object[] { new XAttribute("val", replacementHex) };
            }

            colorElement.ReplaceWith(new XElement(elementName, content));
        }

        private static AhoCorasickMatcher<string> BuildMatcher(IEnumerable<string> patterns)
        {
            var matcher = new AhoCorasickMatcher<string>(StringComparer.OrdinalIgnoreCase);
            var added = false;
            foreach (var pattern in patterns ?? Enumerable.Empty<string>())
            {
                if (string.IsNullOrWhiteSpace(pattern))
                {
                    continue;
                }

                matcher.Add(pattern, pattern);
                added = true;
            }

            if (!added)
            {
                return null;
            }

            matcher.Build();
            return matcher;
        }

        private static IReadOnlyList<PartDescriptor> BuildSlidePartDescriptors(PresentationDocument document)
        {
            var descriptors = new List<PartDescriptor>();
            if (document == null || document.PresentationPart == null || document.PresentationPart.Presentation == null)
            {
                return descriptors;
            }

            var index = 1;
            var slideIds = document.PresentationPart.Presentation.SlideIdList == null
                ? Enumerable.Empty<SlideId>()
                : document.PresentationPart.Presentation.SlideIdList.Elements<SlideId>();
            foreach (var slideId in slideIds)
            {
                if (slideId == null || slideId.RelationshipId == null)
                {
                    continue;
                }

                var slidePart = document.PresentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                if (slidePart == null)
                {
                    continue;
                }

                descriptors.Add(new PartDescriptor(slidePart, index));
                index++;
            }

            return descriptors;
        }

        private static IReadOnlyList<PartDescriptor> BuildFontPartDescriptors(PresentationDocument document)
        {
            var descriptors = new List<PartDescriptor>();
            if (document == null || document.PresentationPart == null)
            {
                return descriptors;
            }

            var sortOrder = 0;
            foreach (var slideMaster in document.PresentationPart.SlideMasterParts
                .OrderBy(x => x.Uri == null ? string.Empty : x.Uri.ToString(), StringComparer.OrdinalIgnoreCase))
            {
                descriptors.Add(new PartDescriptor(slideMaster, sortOrder++));

                foreach (var layout in slideMaster.SlideLayoutParts
                    .OrderBy(x => x.Uri == null ? string.Empty : x.Uri.ToString(), StringComparer.OrdinalIgnoreCase))
                {
                    descriptors.Add(new PartDescriptor(layout, sortOrder++));
                }
            }

            foreach (var slide in BuildSlidePartDescriptors(document))
            {
                descriptors.Add(new PartDescriptor(slide.Part, 1000 + slide.SortOrder));
            }

            if (document.PresentationPart.NotesMasterPart != null)
            {
                descriptors.Add(new PartDescriptor(document.PresentationPart.NotesMasterPart, 3000));
            }

            return descriptors;
        }

        private static IDictionary<string, OpenXmlPart> BuildSlidePartLookup(PresentationDocument document)
        {
            return BuildSlidePartDescriptors(document)
                .GroupBy(x => x.Uri, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(x => x.Key, x => x.First().Part, StringComparer.OrdinalIgnoreCase);
        }

        private static IDictionary<string, OpenXmlPart> BuildFontPartLookup(PresentationDocument document)
        {
            return BuildFontPartDescriptors(document)
                .GroupBy(x => x.Uri, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(x => x.Key, x => x.First().Part, StringComparer.OrdinalIgnoreCase);
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

        private static string BuildInstructionKey(ColorUsageKind usageKind, string hexValue)
        {
            return ((int)usageKind).ToString() + "|" + NormalizeHex(hexValue);
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

        private static string NormalizeHex(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var normalized = value.Trim().TrimStart('#').ToUpperInvariant();
            return normalized.Length == 6 ? normalized : string.Empty;
        }

        private static string NormalizeFontName(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
        }

        private static bool IsDirectColorElement(XElement element)
        {
            if (element == null || element.Name.NamespaceName != DrawingNamespace)
            {
                return false;
            }

            return element.Name.LocalName.Equals("srgbClr", StringComparison.OrdinalIgnoreCase)
                || element.Name.LocalName.Equals("sysClr", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsFontElement(XElement element)
        {
            if (element == null || element.Name.NamespaceName != DrawingNamespace)
            {
                return false;
            }

            return element.Name.LocalName.Equals("latin", StringComparison.OrdinalIgnoreCase)
                || element.Name.LocalName.Equals("ea", StringComparison.OrdinalIgnoreCase)
                || element.Name.LocalName.Equals("cs", StringComparison.OrdinalIgnoreCase)
                || element.Name.LocalName.Equals("sym", StringComparison.OrdinalIgnoreCase)
                || element.Name.LocalName.Equals("buFont", StringComparison.OrdinalIgnoreCase);
        }

        private static ColorUsageKind ResolveUsageKind(XElement colorElement)
        {
            if (colorElement == null)
            {
                return ColorUsageKind.ShapeFill;
            }

            if (colorElement.Ancestors().Any(x => x.Name.NamespaceName == ChartNamespace))
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

        private sealed class XmlWorkItem
        {
            public XmlWorkItem(string partUri, int sortOrder, string xmlText)
            {
                PartUri = partUri ?? string.Empty;
                SortOrder = sortOrder;
                XmlText = xmlText ?? string.Empty;
            }

            public string PartUri { get; }

            public int SortOrder { get; }

            public string XmlText { get; }
        }

        private sealed class XmlMutationResult
        {
            public XmlMutationResult(string partUri, int sortOrder, string xmlText, int changeCount)
            {
                PartUri = partUri ?? string.Empty;
                SortOrder = sortOrder;
                XmlText = xmlText ?? string.Empty;
                ChangeCount = changeCount;
            }

            public string PartUri { get; }

            public int SortOrder { get; }

            public string XmlText { get; }

            public int ChangeCount { get; }
        }

        private sealed class PartDescriptor
        {
            public PartDescriptor(OpenXmlPart part, int sortOrder)
            {
                Part = part;
                SortOrder = sortOrder;
                Uri = part == null || part.Uri == null
                    ? string.Empty
                    : part.Uri.ToString();
            }

            public OpenXmlPart Part { get; }

            public int SortOrder { get; }

            public string Uri { get; }
        }
    }
}
