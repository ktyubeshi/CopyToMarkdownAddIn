namespace CopyFromExcelToMarkdownAddIn
{
    /// <summary>
    /// Represents a table block in Markdown.
    /// </summary>
    public class TableBlock : IMarkdownBlock
    {
        /// <summary>
        /// The table data for this block.
        /// </summary>
        public Table TableData { get; }

        /// <summary>
        /// Creates a new TableBlock with the specified table data.
        /// </summary>
        /// <param name="tableData">The table data for this block.</param>
        public TableBlock(Table tableData)
        {
            TableData = tableData;
        }
    }
}
