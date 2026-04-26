using System;
using System.Collections.Generic;
using System.Linq;

namespace MorphosPowerPointAddIn.Utilities
{
    internal sealed class FontTrie
    {
        private readonly Node _root = new Node();

        public void Add(string fontName)
        {
            if (string.IsNullOrWhiteSpace(fontName))
            {
                return;
            }

            var normalized = fontName.Trim();
            var current = _root;
            foreach (var symbol in normalized)
            {
                var key = Normalize(symbol);
                Node next;
                if (!current.Children.TryGetValue(key, out next))
                {
                    next = new Node();
                    current.Children[key] = next;
                }

                current = next;
                current.PassThroughCount++;
            }

            current.IsTerminal = true;
            current.Value = normalized;
        }

        public IReadOnlyList<string> Search(string prefix, int maximumResults)
        {
            if (maximumResults <= 0)
            {
                return Array.Empty<string>();
            }

            var normalizedPrefix = string.IsNullOrWhiteSpace(prefix)
                ? string.Empty
                : prefix.Trim();

            var current = _root;
            foreach (var symbol in normalizedPrefix)
            {
                Node next;
                if (!current.Children.TryGetValue(Normalize(symbol), out next))
                {
                    return Array.Empty<string>();
                }

                current = next;
            }

            var results = new List<string>(maximumResults);
            Collect(current, results, maximumResults);
            return results;
        }

        private static void Collect(Node node, ICollection<string> results, int maximumResults)
        {
            if (node == null || results == null || results.Count >= maximumResults)
            {
                return;
            }

            if (node.IsTerminal && !string.IsNullOrWhiteSpace(node.Value))
            {
                results.Add(node.Value);
            }

            foreach (var child in node.Children
                .OrderByDescending(x => x.Value.PassThroughCount)
                .ThenBy(x => x.Key))
            {
                Collect(child.Value, results, maximumResults);
                if (results.Count >= maximumResults)
                {
                    return;
                }
            }
        }

        private static char Normalize(char value)
        {
            return char.ToUpperInvariant(value);
        }

        private sealed class Node
        {
            public Node()
            {
                Children = new Dictionary<char, Node>();
            }

            public Dictionary<char, Node> Children { get; }

            public string Value { get; set; }

            public bool IsTerminal { get; set; }

            public int PassThroughCount { get; set; }
        }
    }
}
