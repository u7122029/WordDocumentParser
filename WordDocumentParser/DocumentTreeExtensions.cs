using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace WordDocumentParser
{
    /// <summary>
    /// Extension methods for working with the document tree
    /// </summary>
    public static class DocumentTreeExtensions
    {
        /// <summary>
        /// Finds all nodes matching a predicate
        /// </summary>
        public static IEnumerable<DocumentNode> FindAll(this DocumentNode root, Func<DocumentNode, bool> predicate)
        {
            if (predicate(root))
                yield return root;

            foreach (var child in root.Children)
            {
                foreach (var match in child.FindAll(predicate))
                {
                    yield return match;
                }
            }
        }

        /// <summary>
        /// Finds the first node matching a predicate
        /// </summary>
        public static DocumentNode? FindFirst(this DocumentNode root, Func<DocumentNode, bool> predicate)
        {
            return root.FindAll(predicate).FirstOrDefault();
        }

        /// <summary>
        /// Gets all headings at a specific level
        /// </summary>
        public static IEnumerable<DocumentNode> GetHeadingsAtLevel(this DocumentNode root, int level)
        {
            return root.FindAll(n => n.Type == ContentType.Heading && n.HeadingLevel == level);
        }

        /// <summary>
        /// Gets all headings in the document
        /// </summary>
        public static IEnumerable<DocumentNode> GetAllHeadings(this DocumentNode root)
        {
            return root.FindAll(n => n.Type == ContentType.Heading);
        }

        /// <summary>
        /// Gets all tables in the document
        /// </summary>
        public static IEnumerable<DocumentNode> GetAllTables(this DocumentNode root)
        {
            return root.FindAll(n => n.Type == ContentType.Table);
        }

        /// <summary>
        /// Gets all images in the document
        /// </summary>
        public static IEnumerable<DocumentNode> GetAllImages(this DocumentNode root)
        {
            return root.FindAll(n => n.Type == ContentType.Image);
        }

        /// <summary>
        /// Gets the table of contents as a flat list with indentation info
        /// </summary>
        public static List<(int Level, string Title, DocumentNode Node)> GetTableOfContents(this DocumentNode root)
        {
            return root.GetAllHeadings()
                       .Select(h => (h.HeadingLevel, h.Text, h))
                       .ToList();
        }

        /// <summary>
        /// Gets all text content under a node (recursive)
        /// </summary>
        public static string GetAllText(this DocumentNode node)
        {
            var texts = new List<string>();

            if (!string.IsNullOrEmpty(node.Text) && node.Type != ContentType.Table && node.Type != ContentType.Image)
            {
                texts.Add(node.Text);
            }

            foreach (var child in node.Children)
            {
                texts.Add(child.GetAllText());
            }

            return string.Join("\n", texts.Where(t => !string.IsNullOrWhiteSpace(t)));
        }

        /// <summary>
        /// Gets the path from root to this node
        /// </summary>
        public static List<DocumentNode> GetPath(this DocumentNode node)
        {
            var path = new List<DocumentNode>();
            var current = node;

            while (current != null)
            {
                path.Insert(0, current);
                current = current.Parent;
            }

            return path;
        }

        /// <summary>
        /// Gets the heading path (breadcrumb) for a node
        /// </summary>
        public static string GetHeadingPath(this DocumentNode node, string separator = " > ")
        {
            var headings = node.GetPath()
                               .Where(n => n.Type == ContentType.Heading || n.Type == ContentType.Document)
                               .Select(n => n.Text)
                               .Where(t => !string.IsNullOrEmpty(t));

            return string.Join(separator, headings);
        }

        /// <summary>
        /// Gets siblings of a node
        /// </summary>
        public static IEnumerable<DocumentNode> GetSiblings(this DocumentNode node)
        {
            if (node.Parent == null)
                return Enumerable.Empty<DocumentNode>();

            return node.Parent.Children.Where(c => c != node);
        }

        /// <summary>
        /// Gets the next sibling
        /// </summary>
        public static DocumentNode? GetNextSibling(this DocumentNode node)
        {
            if (node.Parent == null) return null;

            var siblings = node.Parent.Children;
            var index = siblings.IndexOf(node);

            return index >= 0 && index < siblings.Count - 1 ? siblings[index + 1] : null;
        }

        /// <summary>
        /// Gets the previous sibling
        /// </summary>
        public static DocumentNode? GetPreviousSibling(this DocumentNode node)
        {
            if (node.Parent == null) return null;

            var siblings = node.Parent.Children;
            var index = siblings.IndexOf(node);

            return index > 0 ? siblings[index - 1] : null;
        }

        /// <summary>
        /// Counts nodes by type
        /// </summary>
        public static Dictionary<ContentType, int> CountByType(this DocumentNode root)
        {
            var counts = new Dictionary<ContentType, int>();

            foreach (var node in root.FindAll(_ => true))
            {
                if (!counts.ContainsKey(node.Type))
                    counts[node.Type] = 0;
                counts[node.Type]++;
            }

            return counts;
        }

        /// <summary>
        /// Flattens the tree to a list in document order
        /// </summary>
        public static List<DocumentNode> Flatten(this DocumentNode root)
        {
            var result = new List<DocumentNode> { root };

            foreach (var child in root.Children)
            {
                result.AddRange(child.Flatten());
            }

            return result;
        }

        /// <summary>
        /// Gets a section by heading text (case-insensitive partial match)
        /// </summary>
        public static DocumentNode? GetSection(this DocumentNode root, string headingText)
        {
            return root.FindFirst(n =>
                n.Type == ContentType.Heading &&
                n.Text.Contains(headingText, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Extracts TableData from a table node
        /// </summary>
        public static TableData? GetTableData(this DocumentNode tableNode)
        {
            if (tableNode.Type != ContentType.Table)
                return null;

            return tableNode.Metadata.TryGetValue("TableData", out var data)
                ? data as TableData
                : null;
        }

        /// <summary>
        /// Extracts ImageData from an image node
        /// </summary>
        public static ImageData? GetImageData(this DocumentNode imageNode)
        {
            if (imageNode.Type != ContentType.Image)
                return null;

            return imageNode.Metadata.TryGetValue("ImageData", out var data)
                ? data as ImageData
                : null;
        }

        /// <summary>
        /// Saves the document tree to a Word document file (.docx)
        /// </summary>
        /// <param name="root">The root document node</param>
        /// <param name="filePath">The path where the document will be saved</param>
        public static void SaveToFile(this DocumentNode root, string filePath)
        {
            using var writer = new WordDocumentTreeWriter();
            writer.WriteToFile(root, filePath);
        }

        /// <summary>
        /// Saves the document tree to a stream as a Word document (.docx)
        /// </summary>
        /// <param name="root">The root document node</param>
        /// <param name="stream">The stream to write to</param>
        public static void SaveToStream(this DocumentNode root, Stream stream)
        {
            using var writer = new WordDocumentTreeWriter();
            writer.WriteToStream(root, stream);
        }

        /// <summary>
        /// Saves the document tree to a byte array as a Word document (.docx)
        /// </summary>
        /// <param name="root">The root document node</param>
        /// <returns>The document as a byte array</returns>
        public static byte[] ToDocxBytes(this DocumentNode root)
        {
            using var stream = new MemoryStream();
            root.SaveToStream(stream);
            return stream.ToArray();
        }
    }
}
