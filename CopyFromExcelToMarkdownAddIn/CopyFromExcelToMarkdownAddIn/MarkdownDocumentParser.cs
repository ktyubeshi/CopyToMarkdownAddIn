using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace CopyFromExcelToMarkdownAddIn
{
    /// <summary>
    /// Parses a Markdown document into blocks (text blocks and table blocks).
    /// </summary>
    public class MarkdownDocumentParser
    {
        private static readonly Regex TableLinePattern = new Regex(@"^\|.*\|$", RegexOptions.Compiled);

        /// <summary>
        /// Parses the input Markdown text into a list of blocks.
        /// </summary>
        /// <param name="text">The Markdown text to parse.</param>
        /// <returns>A list of IMarkdownBlock instances.</returns>
        public List<IMarkdownBlock> Parse(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return new List<IMarkdownBlock>();
            }

            var blocks = new List<IMarkdownBlock>();
            var tableBuffer = new StringBuilder();
            var isInTableMode = false;

            using (var reader = new StringReader(text))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    // Check if this line is a table row
                    bool isTableLine = IsTableLine(line);

                    if (isTableLine)
                    {
                        // Start or continue table mode
                        isInTableMode = true;
                        tableBuffer.AppendLine(line);
                    }
                    else
                    {
                        // Not a table line
                        if (isInTableMode)
                        {
                            // End of table block - flush the buffer
                            FlushTableBuffer(tableBuffer, blocks);
                            isInTableMode = false;
                        }

                        // Add as text block (unless it's an empty line after a table)
                        if (!string.IsNullOrWhiteSpace(line) || blocks.Count == 0 || !(blocks[blocks.Count - 1] is TableBlock))
                        {
                            blocks.Add(new TextBlock(line));
                        }
                    }
                }

                // Flush any remaining table buffer at the end of the document
                if (isInTableMode && tableBuffer.Length > 0)
                {
                    FlushTableBuffer(tableBuffer, blocks);
                }
            }

            return blocks;
        }

        /// <summary>
        /// Checks if a line is a table row (starts and ends with pipe character).
        /// </summary>
        /// <param name="line">The line to check.</param>
        /// <returns>True if the line is a table row, false otherwise.</returns>
        private bool IsTableLine(string line)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                return false;
            }

            var trimmed = line.Trim();
            return TableLinePattern.IsMatch(trimmed);
        }

        /// <summary>
        /// Flushes the table buffer by parsing it and adding a TableBlock to the blocks list.
        /// </summary>
        /// <param name="tableBuffer">The buffer containing table lines.</param>
        /// <param name="blocks">The list of blocks to add the TableBlock to.</param>
        private void FlushTableBuffer(StringBuilder tableBuffer, List<IMarkdownBlock> blocks)
        {
            if (tableBuffer.Length == 0)
            {
                return;
            }

            try
            {
                var tableText = tableBuffer.ToString();
                var grid = new GridParser().Parse(tableText);
                var table = new TableParser().Parse(grid);
                blocks.Add(new TableBlock(table));
            }
            catch (Exception)
            {
                // If table parsing fails, treat the entire buffer as text
                // This provides a fallback for malformed tables
                var lines = tableBuffer.ToString().Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                foreach (var line in lines)
                {
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        blocks.Add(new TextBlock(line));
                    }
                }
            }
            finally
            {
                tableBuffer.Clear();
            }
        }
    }
}
