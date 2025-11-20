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

        /// <summary>
        /// Converts the block to a Markdown string.
        /// </summary>
        /// <returns>The Markdown representation of the block.</returns>
        public string ToMarkdown()
        {
            if (TableData == null || TableData.Rows.Count == 0)
            {
                return string.Empty;
            }

            var builder = new System.Text.StringBuilder();
            var rows = TableData.Rows;
            var columnCount = rows[0].Count;

            // 1. Header
            builder.Append("|");
            foreach (var cell in rows[0])
            {
                builder.Append(EscapeCellText(cell.Value));
                builder.Append("|");
            }
            builder.AppendLine();

            // 2. Separator
            builder.Append("|");
            foreach (var cell in rows[0])
            {
                switch (cell.Alignment)
                {
                    case Alignment.Left:
                        builder.Append(":---|");
                        break;
                    case Alignment.Center:
                        builder.Append(":-:|");
                        break;
                    case Alignment.Right:
                        builder.Append("---:|");
                        break;
                    default:
                        builder.Append("---|");
                        break;
                }
            }
            builder.AppendLine();

            // 3. Body
            for (int i = 1; i < rows.Count; i++)
            {
                var row = rows[i];
                builder.Append("|");
                // Handle potential row length mismatch if any, though Parser should align it.
                for (int j = 0; j < columnCount; j++)
                {
                    if (j < row.Count)
                    {
                        builder.Append(EscapeCellText(row[j].Value));
                    }
                    builder.Append("|");
                }
                builder.AppendLine();
            }

            return builder.ToString();
        }

        private string EscapeCellText(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            return text.Replace("|", "\\|").Replace("\n", "<br>");
        }
    }
}
