namespace CopyFromExcelToMarkdownAddIn
{
    /// <summary>
    /// Represents a text block in Markdown (heading, paragraph, or plain text line).
    /// </summary>
    public class TextBlock : IMarkdownBlock
    {
        /// <summary>
        /// The text content of this block.
        /// </summary>
        public string Text { get; }

        /// <summary>
        /// Creates a new TextBlock with the specified text.
        /// </summary>
        /// <param name="text">The text content of this block.</param>
        public TextBlock(string text)
        {
            Text = text;
        }

        /// <summary>
        /// Converts the block to a Markdown string.
        /// </summary>
        /// <returns>The Markdown representation of the block.</returns>
        public string ToMarkdown()
        {
            return Text;
        }
    }
}
