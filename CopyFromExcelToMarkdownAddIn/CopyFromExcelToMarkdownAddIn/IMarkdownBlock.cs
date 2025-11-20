namespace CopyFromExcelToMarkdownAddIn
{
    /// <summary>
    /// Base interface for Markdown document blocks.
    /// A block can be either a text block (heading/paragraph) or a table block.
    /// </summary>
    public interface IMarkdownBlock
    {
        /// <summary>
        /// Converts the block to a Markdown string.
        /// </summary>
        /// <returns>The Markdown representation of the block.</returns>
        string ToMarkdown();
    }
}
