using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace CopyFromExcelToMarkdownAddIn
{
    public class RangeParser
    {
        private const int AlignmentLeft = -4131;
        private const int AlignmentCenter = -4108;
        private const int AlignmentRight = -4152;

        public List<IMarkdownBlock> Parse(Range range)
        {
            var blocks = new List<IMarkdownBlock>();
            if (range == null) return blocks;

            int rowCount = range.Rows.Count;
            int colCount = range.Columns.Count;
            TableBlock currentTable = null;

            for (int r = 1; r <= rowCount; r++)
            {
                Range rowRange = range.Rows[r];
                
                // Analyze the row content
                int nonEmptyCount = 0;
                int firstNonEmptyColIndex = -1;
                
                // Check cells in the row
                for (int c = 1; c <= colCount; c++)
                {
                    Range cell = rowRange.Cells[1, c];
                    string text = cell.Text as string;
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        nonEmptyCount++;
                        if (firstNonEmptyColIndex == -1) firstNonEmptyColIndex = c;
                    }
                }

                // Determine if this is a Table Row or Text Row
                // Heuristic: 
                // - If multiple columns have data -> Table.
                // - If it's a single column but we are already in a table -> 
                //   Check if it looks like a continuation? 
                //   Actually, simplified: Multiple cols -> Table. Single col -> Text.
                //   (This means single-column tables are converted to text/list, which is usually fine).
                
                bool isTableRow = (nonEmptyCount > 1) || (currentTable != null && nonEmptyCount > 1);
                // Also consider the case where a table has a row with only 1 cell filled?
                // e.g. | A | B | \n | Value | | 
                // If we treat that as Text, we break the table.
                // So if currentTable != null, and nonEmptyCount > 0, we might want to keep it as table IF it fits?
                // But "Headings" (single cell) should break tables.
                // Let's strictly say: If it looks like a Header (Bold/Large), it breaks table.
                // Otherwise, if we are in a table, we try to stay in it?
                
                // Let's stick to the simple rule first to satisfy the prompt "Handle Headings/Lists".
                // Those are typically distinct from tables.
                
                if (nonEmptyCount > 1)
                {
                    isTableRow = true;
                }
                else if (currentTable != null && nonEmptyCount > 0)
                {
                     // We are in a table, but this row has only 1 cell.
                     // Is it a Heading?
                     Range cell = rowRange.Cells[1, firstNonEmptyColIndex];
                     if (IsHeading(cell))
                     {
                         isTableRow = false;
                     }
                     else
                     {
                         // Treat as sparse table row
                         isTableRow = true;
                     }
                }
                else
                {
                    isTableRow = false;
                }

                if (isTableRow)
                {
                    if (currentTable == null)
                    {
                        currentTable = new TableBlock(new Table());
                        blocks.Add(currentTable);
                    }

                    var tr = new TableRow();
                    for (int c = 1; c <= colCount; c++)
                    {
                        Range cell = rowRange.Cells[1, c];
                        tr.AddCell(new TableCell(cell.Text?.ToString() ?? "", GetAlignment(cell.HorizontalAlignment)));
                    }
                    currentTable.TableData.AddRow(tr);
                }
                else
                {
                    currentTable = null; // End table

                    if (nonEmptyCount == 0) continue;

                    Range cell = rowRange.Cells[1, firstNonEmptyColIndex];
                    string text = cell.Text?.ToString() ?? "";
                    // Normalize newlines
                    text = text.Replace("\n", " "); 

                    if (IsHeading1(cell))
                    {
                        blocks.Add(new TextBlock("# " + text));
                    }
                    else if (IsHeading2(cell))
                    {
                        blocks.Add(new TextBlock("## " + text));
                    }
                    else if (IsHeading3(cell))
                    {
                         blocks.Add(new TextBlock("### " + text));
                    }
                    else if (IsList(cell))
                    {
                        string prefix = "* ";
                        if (IsOrderedList(cell))
                        {
                            // Try to preserve numbering if manual?
                            // Or just use "1. "
                            // If Excel uses "1." manually, we keep it.
                            // If Excel uses auto-numbering... `cell.Text` usually contains it? 
                            // Excel doesn't have "Auto Numbering" in cells like Word. It's manual.
                            // Unless "List" style?
                            // We'll assume unorderd list for IndentLevel logic for now.
                            prefix = "* ";
                        }
                        
                        int indent = (int)cell.IndentLevel;
                        string indentation = new string(' ', indent * 2);
                        
                        // Clean existing bullet chars
                        string cleanText = text.Trim();
                        if (cleanText.StartsWith("•") || cleanText.StartsWith("-") || cleanText.StartsWith("*"))
                        {
                            cleanText = cleanText.Substring(1).TrimStart();
                        }
                        
                        blocks.Add(new TextBlock(indentation + prefix + cleanText));
                    }
                    else
                    {
                        blocks.Add(new TextBlock(text));
                    }
                }
            }

            return blocks;
        }

        private Alignment GetAlignment(dynamic horizontalAlignment)
        {
            // Excel returns 'dynamic' (int or DBNull/null)
            if (!(horizontalAlignment is int)) return Alignment.Undefined;
            int align = (int)horizontalAlignment;

            switch (align)
            {
                case AlignmentLeft: return Alignment.Left;
                case AlignmentCenter: return Alignment.Center;
                case AlignmentRight: return Alignment.Right;
                default: return Alignment.Undefined;
            }
        }

                private bool IsHeading(Range cell)

                {

                    return IsHeading1(cell) || IsHeading2(cell) || IsHeading3(cell);

                }

        

                private bool IsHeading1(Range cell)

                {

                    object size = cell.Font.Size;

                    if (size is double d)

                    {

                        return d >= 18;

                    }

                    return false;

                }

        

                private bool IsHeading2(Range cell)

                {

                    object size = cell.Font.Size;

                    if (size is double d)

                    {

                        return d >= 14 && d < 18;

                    }

                    return false;

                }

        

                private bool IsHeading3(Range cell)

                {

                    object bold = cell.Font.Bold;

                    // Excel returns true/false or DBNull (if mixed)

                    if (bold is bool b)

                    {

                        return b;

                    }

                    return false;

                }

        

                private bool IsList(Range cell)

                {

                    object indent = cell.IndentLevel;

                    if (indent is int i && i > 0) return true;

                    if (indent is double d && d > 0) return true; // Just in case

        

                    string t = cell.Text?.ToString().Trim();

                    if (string.IsNullOrEmpty(t)) return false;

                    return t.StartsWith("•") || t.StartsWith("- ") || t.StartsWith("* ");

                }

        private bool IsOrderedList(Range cell)
        {
            // Not implemented detection of "1." etc yet, relying on manual text.
            return false;
        }
    }
}
