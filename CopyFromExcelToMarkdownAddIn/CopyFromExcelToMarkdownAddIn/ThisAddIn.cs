using System;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace CopyFromExcelToMarkdownAddIn
{
    public partial class ThisAddIn
    {
        /// <summary>
        /// Alignment Undefined
        /// </summary>
        private const int AlignmentUndefined = 1;
        /// <summary>
        /// Alignment Left
        /// </summary>
        private const int AlignmentLeft = -4131;
        /// <summary>
        /// Alignment Center
        /// </summary>
        private const int AlignmentCenter = -4108;
        /// <summary>
        /// Alignment Right
        /// </summary>
        private const int AlignmentRight = -4152;

        /// <summary>
        /// Button in ContextMenu for Cell.
        /// </summary>
        private CommandBarButton _copyToMarkdownButtonForCell;
        /// <summary>
        /// Button in ContextMenu for Table.
        /// </summary>
        private CommandBarButton _copyToMarkDownButtonForTable;

        /// <summary>
        /// Button in ContextMenu for Cell.
        /// </summary>
        private CommandBarButton _copyFromMarkdownButtonForCell;
        /// <summary>
        /// Button in ContextMenu for Table.
        /// </summary>
        private CommandBarButton _copyFromMarkdownButtonForTable;

        /// <summary>
        /// Startup event handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            const string CELL = "Cell";
            const string TABLE = "List Range Popup";

            // Create button in ContextMenu for Cell
            _copyToMarkdownButtonForCell = CreateCopyToMarkdownButton(CELL, "0");
            _copyFromMarkdownButtonForCell = CreateCopyFromMarkdownButton(CELL, "1");
            // Create button in ContextMenu for Table
            _copyToMarkDownButtonForTable = CreateCopyToMarkdownButton(TABLE, "2");
            _copyFromMarkdownButtonForTable = CreateCopyFromMarkdownButton(TABLE, "3");
        }

        private CommandBarButton CreateCopyToMarkdownButton(string commandBarsKey, string tag)
        {
            var copyToMarkdownButton = (CommandBarButton)Application.CommandBars[commandBarsKey].Controls.Add(MsoControlType.msoControlButton, missing, missing, 1, true);
            copyToMarkdownButton.Style = MsoButtonStyle.msoButtonCaption;
            copyToMarkdownButton.Caption = "Copy to Markdown";
            copyToMarkdownButton.Tag = tag;
            copyToMarkdownButton.Click += CopyToMarkdown;
            return copyToMarkdownButton;
        }
        private CommandBarButton CreateCopyFromMarkdownButton(string commandBarsKey, string tag)
        {
            var copyFromMarkdownButton = (CommandBarButton)Application.CommandBars[commandBarsKey].Controls.Add(MsoControlType.msoControlButton, missing, missing, 2, true);
            copyFromMarkdownButton.Style = MsoButtonStyle.msoButtonCaption;
            copyFromMarkdownButton.Caption = "Paste from Markdown";
            copyFromMarkdownButton.Tag = tag;
            copyFromMarkdownButton.Click += CopyFromMarkdown;
            return copyFromMarkdownButton;
        }

        private void CopyFromMarkdown(CommandBarButton ctrl, ref bool canceldefault)
        {
            var text = Clipboard.GetText();
            if(string.IsNullOrEmpty(text))
                return;

            var range = Application.Selection as Range;
            if (range == null)
            {
                MessageBox.Show(Properties.Resources.UnselectedErrorMessage);
                return;
            }

            // Parse the Markdown document into blocks
            var blocks = new MarkdownDocumentParser().Parse(text);
            var activeSheet = (Worksheet)Application.ActiveSheet;
            var currentRow = range.Row;

            // Process each block
            foreach (var block in blocks)
            {
                if (block is TextBlock textBlock)
                {
                    // Write text block to a single cell in the left column
                    WriteTextBlock(textBlock, activeSheet, currentRow, range.Column);
                    currentRow++;
                }
                else if (block is TableBlock tableBlock)
                {
                    // Write table block to cells
                    var rowsWritten = WriteTableBlock(tableBlock, activeSheet, currentRow, range.Column);
                    currentRow += rowsWritten;
                }
            }
        }

        /// <summary>
        /// Writes a text block to a single cell.
        /// </summary>
        /// <param name="textBlock">The text block to write.</param>
        /// <param name="sheet">The worksheet to write to.</param>
        /// <param name="row">The row index to write to.</param>
        /// <param name="column">The column index to write to.</param>
        private void WriteTextBlock(TextBlock textBlock, Worksheet sheet, int row, int column)
        {
            var cell = (Range)sheet.Cells[row, column];
            var text = textBlock.Text;

            // Check if the text contains Markdown formatting markers
            if (!System.Text.RegularExpressions.Regex.IsMatch(text, @"(\*|_|~|`)"))
            {
                cell.Value2 = text;
            }
            else
            {
                // Parse Markdown formatting
                var segments = MarkdownInlineParser.Parse(text);

                // 1. First, set the clean text (without Markdown markers) to the cell
                var cleanTextBuilder = new StringBuilder();
                foreach (var seg in segments)
                {
                    cleanTextBuilder.Append(seg.Text);
                }
                cell.Value2 = cleanTextBuilder.ToString();

                // 2. Apply formatting to each segment
                int currentIndex = 1; // Excel's Characters indexing is 1-based
                foreach (var seg in segments)
                {
                    int len = seg.Text.Length;
                    if (len > 0)
                    {
                        var chars = cell.Characters[currentIndex, len];
                        if (seg.IsBold)
                        {
                            chars.Font.Bold = true;
                        }
                        if (seg.IsItalic)
                        {
                            chars.Font.Italic = true;
                        }
                        if (seg.IsStrikeThrough)
                        {
                            chars.Font.Strikethrough = true;
                        }

                        currentIndex += len;
                    }
                }
            }
        }

        /// <summary>
        /// Writes a table block to cells.
        /// </summary>
        /// <param name="tableBlock">The table block to write.</param>
        /// <param name="sheet">The worksheet to write to.</param>
        /// <param name="startRow">The starting row index.</param>
        /// <param name="startColumn">The starting column index.</param>
        /// <returns>The number of rows written.</returns>
        private int WriteTableBlock(TableBlock tableBlock, Worksheet sheet, int startRow, int startColumn)
        {
            var table = tableBlock.TableData;

            for (var i = 0; i < table.Rows.Count; i++)
            {
                var row = table.Rows[i];
                for (var j = 0; j < row.Count; j++)
                {
                    var cell = row[j];
                    var activeSheetCell = (Range)sheet.Cells[startRow + i, startColumn + j];

                    // Replace <br> tags with newlines
                    var baseText = cell.Value.Replace("<br>", "\n").Replace("<br/>", "\n");

                    // Check if the text contains Markdown formatting markers
                    // If not, skip parsing for better performance
                    if (!System.Text.RegularExpressions.Regex.IsMatch(baseText, @"(\*|_|~|`)"))
                    {
                        activeSheetCell.Value2 = baseText;
                    }
                    else
                    {
                        // Parse Markdown formatting
                        var segments = MarkdownInlineParser.Parse(baseText);

                        // 1. First, set the clean text (without Markdown markers) to the cell
                        var cleanTextBuilder = new StringBuilder();
                        foreach (var seg in segments)
                        {
                            cleanTextBuilder.Append(seg.Text);
                        }
                        activeSheetCell.Value2 = cleanTextBuilder.ToString();

                        // 2. Apply formatting to each segment
                        int currentIndex = 1; // Excel's Characters indexing is 1-based
                        foreach (var seg in segments)
                        {
                            int len = seg.Text.Length;
                            if (len > 0)
                            {
                                var chars = activeSheetCell.Characters[currentIndex, len];
                                if (seg.IsBold)
                                {
                                    chars.Font.Bold = true;
                                }
                                if (seg.IsItalic)
                                {
                                    chars.Font.Italic = true;
                                }
                                if (seg.IsStrikeThrough)
                                {
                                    chars.Font.Strikethrough = true;
                                }

                                currentIndex += len;
                            }
                        }
                    }

                    // Set cell alignment
                    switch (cell.Alignment)
                    {
                        case Alignment.Undefined:
                            activeSheetCell.HorizontalAlignment = AlignmentUndefined;
                            break;
                        case Alignment.Left:
                            activeSheetCell.HorizontalAlignment = AlignmentLeft;
                            break;
                        case Alignment.Center:
                            activeSheetCell.HorizontalAlignment = AlignmentCenter;
                            break;
                        case Alignment.Right:
                            activeSheetCell.HorizontalAlignment = AlignmentRight;
                            break;
                        default:
                            throw new ArgumentOutOfRangeException();
                    }
                }
            }

            return table.Rows.Count;
        }

        /// <summary>
        /// Shutdown event handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        /// <summary>
        /// CopyToMarkdown
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="cancelDefault"></param>
        private void CopyToMarkdown(CommandBarButton ctrl, ref bool cancelDefault)
        {
            var range = Application.Selection as Range;
            if (range == null)
            {
                MessageBox.Show(Properties.Resources.UnselectedErrorMessage);
                return;
            }

            try
            {
                var parser = new RangeParser();
                var blocks = parser.Parse(range);

                var sb = new StringBuilder();
                foreach (var block in blocks)
                {
                    sb.AppendLine(block.ToMarkdown());
                    sb.AppendLine(); // Ensure separation between blocks
                }

                if (sb.Length > 0)
                {
                    Clipboard.SetText(sb.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error copying to Markdown: " + ex.Message);
            }
        }


        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
        
        #endregion
    }
}
