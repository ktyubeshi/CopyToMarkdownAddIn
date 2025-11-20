using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace CopyFromExcelToMarkdownAddIn
{
    /// <summary>
    /// Represents a text segment with style information.
    /// </summary>
    public class StyledTextSegment
    {
        public string Text { get; set; }
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool IsStrikeThrough { get; set; }

        public StyledTextSegment(string text)
        {
            Text = text;
            IsBold = false;
            IsItalic = false;
            IsStrikeThrough = false;
        }
    }

    /// <summary>
    /// Parses Markdown inline formatting (bold, italic, strikethrough) into styled text segments.
    /// </summary>
    public class MarkdownInlineParser
    {
        /// <summary>
        /// Parses Markdown text and returns a list of styled text segments.
        /// Supports: **bold**, *italic*, ~~strikethrough~~
        /// </summary>
        /// <param name="markdown">The Markdown text to parse</param>
        /// <returns>List of styled text segments</returns>
        public static List<StyledTextSegment> Parse(string markdown)
        {
            var segments = new List<StyledTextSegment>();
            if (string.IsNullOrEmpty(markdown))
            {
                return segments;
            }

            int position = 0;
            while (position < markdown.Length)
            {
                // Try to match markdown patterns in order of precedence
                // Check for bold first (** or __)
                if (TryMatchPattern(markdown, position, @"^\*\*(.+?)\*\*", out var boldMatch))
                {
                    var segment = new StyledTextSegment(boldMatch.Groups[1].Value)
                    {
                        IsBold = true
                    };
                    segments.Add(segment);
                    position += boldMatch.Length;
                }
                else if (TryMatchPattern(markdown, position, @"^__(.+?)__", out var boldMatch2))
                {
                    var segment = new StyledTextSegment(boldMatch2.Groups[1].Value)
                    {
                        IsBold = true
                    };
                    segments.Add(segment);
                    position += boldMatch2.Length;
                }
                // Check for strikethrough (~~)
                else if (TryMatchPattern(markdown, position, @"^~~(.+?)~~", out var strikeMatch))
                {
                    var segment = new StyledTextSegment(strikeMatch.Groups[1].Value)
                    {
                        IsStrikeThrough = true
                    };
                    segments.Add(segment);
                    position += strikeMatch.Length;
                }
                // Check for italic (* or _)
                else if (TryMatchPattern(markdown, position, @"^\*(.+?)\*", out var italicMatch))
                {
                    var segment = new StyledTextSegment(italicMatch.Groups[1].Value)
                    {
                        IsItalic = true
                    };
                    segments.Add(segment);
                    position += italicMatch.Length;
                }
                else if (TryMatchPattern(markdown, position, @"^_(.+?)_", out var italicMatch2))
                {
                    var segment = new StyledTextSegment(italicMatch2.Groups[1].Value)
                    {
                        IsItalic = true
                    };
                    segments.Add(segment);
                    position += italicMatch2.Length;
                }
                // Check for inline code (`)
                else if (TryMatchPattern(markdown, position, @"^`(.+?)`", out var codeMatch))
                {
                    // For code, we just add plain text (Excel doesn't have inline code styling)
                    var segment = new StyledTextSegment(codeMatch.Groups[1].Value);
                    segments.Add(segment);
                    position += codeMatch.Length;
                }
                else
                {
                    // No markdown pattern found, consume one character as plain text
                    var plainText = new StringBuilder();
                    plainText.Append(markdown[position]);
                    position++;

                    // Continue consuming characters until we hit a potential markdown marker
                    while (position < markdown.Length && !IsMarkdownMarker(markdown[position]))
                    {
                        plainText.Append(markdown[position]);
                        position++;
                    }

                    segments.Add(new StyledTextSegment(plainText.ToString()));
                }
            }

            // Merge consecutive plain text segments
            return MergeConsecutivePlainSegments(segments);
        }

        /// <summary>
        /// Tries to match a regex pattern at the current position.
        /// </summary>
        private static bool TryMatchPattern(string text, int position, string pattern, out Match match)
        {
            var substring = text.Substring(position);
            match = Regex.Match(substring, pattern);
            return match.Success;
        }

        /// <summary>
        /// Checks if a character is a potential Markdown marker.
        /// </summary>
        private static bool IsMarkdownMarker(char c)
        {
            return c == '*' || c == '_' || c == '~' || c == '`';
        }

        /// <summary>
        /// Merges consecutive segments that have no formatting.
        /// </summary>
        private static List<StyledTextSegment> MergeConsecutivePlainSegments(List<StyledTextSegment> segments)
        {
            var merged = new List<StyledTextSegment>();
            StringBuilder plainTextBuffer = null;

            foreach (var segment in segments)
            {
                if (!segment.IsBold && !segment.IsItalic && !segment.IsStrikeThrough)
                {
                    // Plain text segment
                    if (plainTextBuffer == null)
                    {
                        plainTextBuffer = new StringBuilder();
                    }
                    plainTextBuffer.Append(segment.Text);
                }
                else
                {
                    // Styled segment
                    if (plainTextBuffer != null)
                    {
                        merged.Add(new StyledTextSegment(plainTextBuffer.ToString()));
                        plainTextBuffer = null;
                    }
                    merged.Add(segment);
                }
            }

            // Add remaining plain text if any
            if (plainTextBuffer != null)
            {
                merged.Add(new StyledTextSegment(plainTextBuffer.ToString()));
            }

            return merged;
        }
    }
}
