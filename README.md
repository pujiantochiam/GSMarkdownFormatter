This Google Apps Script code enhances Google Sheets by adding a "Markdown Tools" menu, which provides two main functionalities: importing formatted content from Markdown and a powerful "Find, Replace & Format" tool.

### Key Functionalities:

**1. Markdown Conversion:**
*   Users can paste Markdown text or Markdown tables into dialog boxes.
*   The script parses various Markdown syntax, including:
    *   Headers (H1, H2, H3) with specific font sizes and colors.
    *   Bold, Italic, Bold+Italic, Strikethrough, and Inline Code with distinct styling.
    *   Special formatting for text enclosed in `<u>` tags (maroon color) and square brackets `[]` (dark blue color).
    *   A specific style is applied to the keywords "Example:" and "Examples:".
*   The script intelligently handles line breaks and lists to create a clean layout within a single cell.
*   For tables, it correctly parses the data, applies a blue top border to the data range, colors the first data row's background for visibility, and sets text wrapping and vertical alignment for the entire table.

**2. Find, Replace & Format Tool:**
*   This feature is accessible through a user-friendly sidebar.
*   Users can search for specific words or exact phrases within the current selection or the entire sheet.
*   It allows for case-sensitive searches and can be configured to match whole words only.
*   **Formatting:** Users can apply various formatting options—such as bold, italics, font color, and font size—to all found instances of the text. This is a "non-destructive" operation, meaning it preserves any pre-existing formatting in the cells.
*   **Replacing:** Optionally, users can replace the found text with new text. When replacing, the specified formatting is applied to the new text.

### How It Works:

*   **`onOpen(e)`**: This function runs automatically when the spreadsheet is opened, creating the "Markdown Tools" menu in the Google Sheets UI.
*   **`showMarkdownInputDialog()` & `showMarkdownTableDialog()`**: These functions display HTML dialogs (`MarkdownInputDialog.html` and `MarkdownTableDialog.html`) where the user can paste their Markdown content.
*   **`processPastedMarkdown(markdownText)` & `processPastedMarkdownTable(markdownTableText)`**: These are the backend functions that take the text from the dialogs, use the `convertMarkdownToRichText` and `convertMarkdownCellToRichText` functions for parsing, and then insert the formatted content into the active cell.
*   **`showSidebar()`**: This opens the "Find, Replace & Format" tool's interface, which is defined in `Sidebar.html`.
*   **`processFindAndFormat(options)`**: This is the core function for the find and replace tool. It receives the user's choices from the sidebar and iterates through the selected range of cells to apply the requested changes. It cleverly distinguishes between "format only" and "replace & format" to either preserve or rebuild the cell's rich text content as needed.

### Technical Details:

*   The script uses the `SpreadsheetApp`, `HtmlService`, and `RichTextValueBuilder` services from Google Apps Script to interact with the spreadsheet, create user interfaces, and build complex text formatting.
*   It defines a set of constant text styles (`BOLD_STYLE`, `H1_STYLE`, etc.) for consistency and easy modification.
*   A `DEBUG_MODE` flag is available for developers to get more detailed logs when troubleshooting.
*   The script includes helper functions like `escapeRegExp` to ensure that user input is safely handled within regular expressions.

### Copyright
(c) 2025 Pujianto Chiam
