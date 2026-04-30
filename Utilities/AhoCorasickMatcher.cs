using System;
using System.Collections.Generic;
using System.Linq;

namespace MorphosPowerPointAddIn.Utilities
{
    internal sealed class AhoCorasickMatcher<TValue>
    {
        private readonly Node _root;
        private readonly IEqualityComparer<TValue> _valueComparer;
        
        // Expose to inner node logic if we need to pass the comparer. Since Node is an inner class, we'll pass the comparer into its constructor.
        private bool _built;

        public AhoCorasickMatcher()
            : this(EqualityComparer<TValue>.Default)
        {
        }

        public AhoCorasickMatcher(IEqualityComparer<TValue> valueComparer)
        {
            _valueComparer = valueComparer ?? EqualityComparer<TValue>.Default;
            _root = new Node(_valueComparer);
        }

        public void Add(string pattern, TValue value)
        {
            if (string.IsNullOrWhiteSpace(pattern))
            {
                return;
            }

            var current = _root;
            foreach (var symbol in pattern.Trim())
            {
                var normalized = Normalize(symbol);
                Node next;
                if (!current.Children.TryGetValue(normalized, out next))
                {
                    next = new Node(_valueComparer);
                    current.Children[normalized] = next;
                }

                current = next;
            }

            current.Outputs.Add(value);
            _built = false;
        }

        public void Build()
        {
            var queue = new Queue<Node>();
            foreach (var child in _root.Children.Values)
            {
                child.Failure = _root;
                queue.Enqueue(child);
            }

            while (queue.Count > 0)
            {
                var current = queue.Dequeue();
                foreach (var pair in current.Children)
                {
                    var transition = pair.Key;
                    var target = pair.Value;
                    var fallback = current.Failure;

                    while (fallback != null && !fallback.Children.ContainsKey(transition))
                    {
                        fallback = fallback.Failure;
                    }

                    target.Failure = fallback == null
                        ? _root
                        : fallback.Children[transition];

                    foreach (var output in target.Failure.Outputs)
                    {
                        target.Outputs.Add(output);
                    }

                    queue.Enqueue(target);
                }
            }

            _built = true;
        }

        public bool Matches(string text)
        {
            return Find(text).Count > 0;
        }

        public IReadOnlyList<TValue> Find(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return Array.Empty<TValue>();
            }

            if (!_built)
            {
                Build();
            }

            var values = new HashSet<TValue>(_valueComparer);
            var current = _root;
            foreach (var symbol in text)
            {
                var normalized = Normalize(symbol);
                while (current != _root && !current.Children.ContainsKey(normalized))
                {
                    current = current.Failure ?? _root;
                }

                Node next;
                if (current.Children.TryGetValue(normalized, out next))
                {
                    current = next;
                }

                foreach (var output in current.Outputs)
                {
                    values.Add(output);
                }
            }

            return values.ToList();
        }

        private static char Normalize(char value)
        {
            return char.ToUpperInvariant(value);
        }

        private sealed class Node
        {
            public Node(IEqualityComparer<TValue> comparer)
            {
                Children = new Dictionary<char, Node>();
                Outputs = new HashSet<TValue>(comparer);
            }

            public Dictionary<char, Node> Children { get; }

            public HashSet<TValue> Outputs { get; }

            public Node Failure { get; set; }
        }
    }
}
