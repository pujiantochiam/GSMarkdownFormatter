/**
 * @OnlyCurrentDoc
 */

// --- Configuration ---
const DEBUG_MODE = false; // Set to true for detailed console logging

// --- Styles (Updated) ---
const BOLD_STYLE = SpreadsheetApp.newTextStyle().setBold(true).build();
// Update Italic Style
const ITALIC_STYLE = SpreadsheetApp.newTextStyle()
    .setItalic(true)
    .setForegroundColor("#2F3840") // <-- Added Italic Color
    .build();
const STRIKETHROUGH_STYLE = SpreadsheetApp.newTextStyle().setStrikethrough(true).build();
const CODE_STYLE = SpreadsheetApp.newTextStyle().setFontFamily("Consolas").setForegroundColor("#4a4a4a").build();
// Update Bold+Italic to include italic color
const BOLD_ITALIC_STYLE = SpreadsheetApp.newTextStyle()
    .setBold(true)
    .setItalic(true)
    .setForegroundColor("#2F3840") // <-- Added Italic Color here too
    .build();
const DARK_BLUE_STYLE = SpreadsheetApp.newTextStyle().setForegroundColor("#00008B").build();
const MAROON_STYLE = SpreadsheetApp.newTextStyle().setForegroundColor("#800000").build(); // <-- NEW Maroon Style

// Update Header Styles
const H1_STYLE = SpreadsheetApp.newTextStyle().setBold(true).setFontSize(14).build(); // No change requested
const H2_STYLE = SpreadsheetApp.newTextStyle()
    .setBold(true)
    .setFontSize(12)
    .setForegroundColor("#5D04D9") // <-- Added H2 Color
    .build();
const H3_STYLE = SpreadsheetApp.newTextStyle()
    .setBold(true)
    .setFontSize(11)
    .setForegroundColor("#260E59") // <-- Added H3 Color
    .build();

// Default Style
const DEFAULT_STYLE = SpreadsheetApp.newTextStyle().setFontSize(10).build();

// Specific Style for "Examples:"
const EXAMPLES_STYLE_COLOR = "#73320D";

// --- NEW Table Border Constants ---
const TABLE_TOP_BORDER_COLOR = "#072BF2"; // Changed to Blue
const TABLE_TOP_BORDER_STYLE = SpreadsheetApp.BorderStyle.SOLID_THICK; // Closest to 1.5pt
const TABLE_HEADER_BACKGROUND_COLOR = "#B3BDF2"; // NEW - Header background color

/**
 * Creates the main menu, combining Markdown Tools and the Find & Format tool.
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Markdown Tools');

  // --- Items from your original Markdown script ---
  menu.addItem('Paste Markdown Here', 'showMarkdownInputDialog');
  menu.addItem('Paste Markdown Table Here', 'showMarkdownTableDialog');
  
  menu.addSeparator(); // Adds a visual line to separate the tools

  // --- NEW: Submenu for the Find, Replace & Format tool ---
  const findFormatSubMenu = ui.createMenu('Find, Replace & Format');
  findFormatSubMenu.addItem('Start Tool', 'showSidebar'); // Points to the new function
  
  menu.addSubMenu(findFormatSubMenu);
  
  menu.addToUi();
}
/**
 * Shows an HTML Service dialog with a textarea for pasting Markdown.
 */
function showMarkdownInputDialog() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('MarkdownInputDialog')
      .setWidth(450) // Adjust width as needed
      .setHeight(350); // Adjust height as needed
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Paste Markdown');
}

function showPasteMarkdownDialog() {
  const ui = SpreadsheetApp.getUi();
  const activeCell = SpreadsheetApp.getActiveRange()?.getCell(1, 1);

  if (!activeCell) {
    ui.alert('Error', 'Please select a cell first.', ui.ButtonSet.OK);
    return;
  }

  const response = ui.prompt(
      'Paste Markdown',
      'Paste your Markdown text below (supports H1-H3, Bold, Italic, Strikethrough, Inline Code, Lists (-/*)):',
      ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const markdownText = response.getResponseText();
    if (markdownText) {
      try {
        activeCell.clearFormat().clearContent();
        const richText = convertMarkdownToRichText(markdownText.trim());
        activeCell.setRichTextValue(richText);
        activeCell.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
        if (DEBUG_MODE) Logger.log("Successfully applied formatting.");
      } catch (error) {
         Logger.log("Error converting Markdown: " + error + "\nStack:\n" + error.stack);
         ui.alert('Error', 'Could not convert Markdown:\n' + error.message, ui.ButtonSet.OK);
         activeCell.setValue(markdownText); // Fallback
      }
    } else {
       if (DEBUG_MODE) Logger.log("User provided empty input.");
       activeCell.clearContent();
    }
  } else {
    if (DEBUG_MODE) Logger.log("User cancelled the prompt.");
  }
}

/**
 * Receives the Markdown text from the HTML dialog, processes it,
 * and pastes the formatted result into the active cell.
 * Called by google.script.run from the client-side HTML JavaScript.
 * @param {string} markdownText The Markdown text pasted by the user.
 */
function processPastedMarkdown(markdownText) {
  const ui = SpreadsheetApp.getUi(); // Get UI for potential errors
  const activeCell = SpreadsheetApp.getActiveRange()?.getCell(1, 1);

  if (!activeCell) {
     // This error needs to be communicated back or handled differently,
     // as we can't easily show a UI alert from here after the dialog closes.
     // Throwing an error will trigger the onFailureHandler in HTML.
     throw new Error("No active cell selected. Please select a cell before pasting.");
     // return; // Or just return if throwing feels too harsh
  }

  if (markdownText) {
    try {
      // --- Call the existing conversion function ---
      Logger.log("Received text from HTML Dialog. Length: " + markdownText.length);
      Logger.log("Contains newline? " + markdownText.includes('\n'));

      activeCell.clearFormat().clearContent();
      const richText = convertMarkdownToRichText(markdownText.trim()); // Use the *same* convert function
      activeCell.setRichTextValue(richText);
      activeCell.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      if (DEBUG_MODE) Logger.log("Successfully applied formatting from HTML Dialog.");
      // No need to return anything on success, the successHandler closes the dialog.
    } catch (error) {
       Logger.log("Error converting Markdown from HTML Dialog: " + error + "\nStack:\n" + error.stack);
       // Throw the error message back to the HTML onFailureHandler
       throw new Error('Could not convert Markdown: ' + error.message);
    }
  } else {
     if (DEBUG_MODE) Logger.log("User provided empty input via HTML Dialog.");
     activeCell.clearContent();
     // Optionally throw an error if empty input is invalid
     // throw new Error("Input cannot be empty.");
  }
}

/**
 * Converts Markdown to RichTextValue using a refined "Collect then Apply" strategy.
 * Applies base line styles first, then specific inline style overrides.
 * Includes custom colors for H2, H3, Italics, and "Examples:".
 *
 * @param {string} markdown The Markdown text to convert.
 * @return {RichTextValue} The formatted RichTextValue object.
 */
function convertMarkdownToRichText(markdown) {
    // --- Initial checks and setup (same as before) ---
    if (!markdown) { return SpreadsheetApp.newRichTextValue().setText("").build(); }
    if (DEBUG_MODE) Logger.log("Starting conversion (Refined Collect & Apply + Colors) for:\n" + markdown);
    const lines = markdown.split('\n');
    let finalText = "";
    const allStyleRuns = [];
    let isFirstLine = true;
    let wasPreviousLineBlock = false;
    let wasPreviousLineList = false;

    for (const line of lines) {
        const trimmedLine = line.trim();
        let lineContent = line;
        let baseStyle = DEFAULT_STYLE;
        let linePrefix = "";
        let isList = false, isHeader = false, isParagraph = false, isEmptyLine = false;

        // --- 1. Determine Line Type, Content, Base Style, and Prefix (Same as before) ---
        if (trimmedLine.startsWith('### ')) { baseStyle = H3_STYLE; lineContent = trimmedLine.substring(4); isHeader = true; }
        else if (trimmedLine.startsWith('## ')) { baseStyle = H2_STYLE; lineContent = trimmedLine.substring(3); isHeader = true; }
        else if (trimmedLine.startsWith('# ')) { baseStyle = H1_STYLE; lineContent = trimmedLine.substring(2); isHeader = true; }
        else if (trimmedLine.startsWith('- ') || trimmedLine.startsWith('* ')) { linePrefix = "•  "; lineContent = trimmedLine.substring(2); baseStyle = DEFAULT_STYLE; isList = true; }
        else if (trimmedLine === "") { isEmptyLine = true; lineContent = ""; }
        else { baseStyle = DEFAULT_STYLE; lineContent = trimmedLine; isParagraph = true; }

        // --- 2. Calculate and Prepend Necessary Newlines (Same as before) ---
        let newlinesToAdd = "";
        const isCurrentLineBlock = isHeader || isList || isParagraph;
        if (!isFirstLine) { /* ... same newline logic ... */
            newlinesToAdd += "\n";
            if (isCurrentLineBlock && wasPreviousLineBlock && !(isList && wasPreviousLineList)) { newlinesToAdd += "\n"; }
            else if (isList && !wasPreviousLineList && wasPreviousLineBlock){ newlinesToAdd += "\n"; }
            else if (isHeader && wasPreviousLineBlock && !wasPreviousLineList) { newlinesToAdd += "\n"; }
         }

        // --- 3. Store Current Position & Append Newlines/Prefix (Same as before) ---
        const lineOverallStartIndex = finalText.length;
        finalText += newlinesToAdd;
        const linePrefixStartIndex = finalText.length;
        finalText += linePrefix;
        const lineContentStartIndex = finalText.length;

        // --- 4. Parse Inline Content & Append its text (Same as before) ---
        const inlineRuns = parseInlineMarkdown(lineContent);
        let lineContentOnlyText = "";
        for(const run of inlineRuns) { lineContentOnlyText += run.text; }
        finalText += lineContentOnlyText;
        const lineContentEndIndex = finalText.length;

        // --- 5. Collect Style Runs (Base first, then Inline Overrides with COLOR logic) ---

        // --- Add BASE Style Run (Same as before) ---
        if (linePrefixStartIndex < lineContentEndIndex) {
             allStyleRuns.push({
                 start: linePrefixStartIndex,
                 end: lineContentEndIndex,
                 style: baseStyle // Base style (includes H2/H3 colors now)
             });
             // if (DEBUG_MODE) Logger.log(...) // Optional logging
        }

        // --- Add INLINE Style Runs (Overrides - WITH COLOR LOGIC) ---
        let currentContentPos = 0;
        for (const run of inlineRuns) {
            const runLength = run.text.length;
            const runTextTrimmedLower = run.text.trim().toLowerCase(); // For checking "examples:"
            const inlineRunStart = lineContentStartIndex + currentContentPos;
            const inlineRunEnd = inlineRunStart + runLength;

            let specificRunStyle = null; // Style for this specific inline segment

            // --- Check for "Examples:"/"Example:" ---
            if (runTextTrimmedLower === 'example:' || runTextTrimmedLower === 'examples:') {
                 // Apply "Examples" color, try to preserve existing style (like bold)
                 const existingStyle = run.style ? run.style : baseStyle; // Inherit from run or line base
                 const examplesSpecificStyle = existingStyle.copy(); // Start with existing style
                 examplesSpecificStyle.setForegroundColor(EXAMPLES_STYLE_COLOR); // Apply specific color
                 specificRunStyle = examplesSpecificStyle.build();
                 if (DEBUG_MODE) Logger.log(`Applied EXAMPLES color to [${inlineRunStart}, ${inlineRunEnd})`);

            // --- Check for other inline styles (Italic now includes color) ---
            } else if (run.style) {
                 // Use the predefined style (e.g., ITALIC_STYLE now has color)
                 specificRunStyle = run.style;
                 // We already defined ITALIC_STYLE and BOLD_ITALIC_STYLE with the color,
                 // so no extra merging needed here for that specifically.
            }

            // --- Add the specific style run if one was determined ---
            if (specificRunStyle && inlineRunStart < inlineRunEnd) {
                 allStyleRuns.push({
                    start: inlineRunStart,
                    end: inlineRunEnd,
                    style: specificRunStyle // This will override the base style for this range
                 });
                 // if (DEBUG_MODE) Logger.log(...) // Optional logging
            }
            currentContentPos += runLength;
        } // End inline runs loop


        // --- 6. Update State for Next Iteration (Same as before) ---
        isFirstLine = false;
        wasPreviousLineBlock = isCurrentLineBlock;
        wasPreviousLineList = isList;
        if(isEmptyLine) { wasPreviousLineBlock = false; wasPreviousLineList = false; }
        // if (DEBUG_MODE) Logger.log(...) // Optional logging

    } // End loop through lines

    // --- 7. Build the RichTextValue and Apply All Styles (Same as before) ---
    if (DEBUG_MODE) { /* ... logging ... */ }
    if (finalText.length === 0) { /* ... handle empty ... */ }

    const builder = SpreadsheetApp.newRichTextValue();
    try {
        builder.setText(finalText); // Set text ONCE

        // Apply all collected styles
        for (const run of allStyleRuns) {
            if (run.start < run.end && run.start < finalText.length && run.end <= finalText.length) {
               try { builder.setTextStyle(run.start, run.end, run.style); }
               catch (eInner) { Logger.log(`ERROR applying style run [${run.start}, ${run.end}) Inner Error: ${eInner}`); }
            } else {
               // if (DEBUG_MODE) Logger.log(...) // Optional logging
            }
        }
    } catch(eOuter) { /* ... error handling ... */
        Logger.log(`FATAL ERROR building RichTextValue: ${eOuter}\nStack: ${eOuter.stack}`);
        return SpreadsheetApp.newRichTextValue().setText(finalText).build(); // Fallback
    }

    return builder.build();
}

/**
 * Parses a single line of text for inline Markdown, <u>, and [...] elements recursively.
 * Returns an array of objects: { text: string, style: RichTextStyle | null }
 * Handles nesting like **[text]** or [<u>text</u>].
 * Converts <u>text</u> to MAROON_STYLE text.
 * Converts [text] to DARK_BLUE_STYLE text.
 * @param {string} textLine The text content to parse.
 * @return {Array<{text: string, style: RichTextStyle | null}>} Segments of text with optional styling.
 */
function parseInlineMarkdown(textLine) {
    const finalSegments = [];
    let currentIndex = 0;

    // Regex Explanation: Added |(\[)(.+?)(\]) for brackets (groups 14, 15, 16)
    // Adjusted plain text [^...] and single char [...] to include escaped brackets \[ and \]
    // Group indices shifted for plain text (17) and single char (18)
    const regex = /(\`+)(.+?)\1|(\*\*\*|___)(.+?)\3|(\*\*|__)(.+?)\5|(\*|_)(.+?)\7|(~~)(.+?)\9|(<u>)(.+?)(<\/u>)|(\[)(.+?)(\])|([^`*_~<>\[\]]+)|([`*_~<>\[\]])/gi; // Added [ and ] to char classes

    let match;
    while ((match = regex.exec(textLine)) !== null) {
        const fullMatch = match[0];
        const matchIndex = match.index;

        // --- 1. Add preceding plain text ---
        if (matchIndex > currentIndex) {
            finalSegments.push({ text: textLine.substring(currentIndex, matchIndex), style: null });
        }

        let styleToApply = null;
        let contentToParse = "";
        let processRecursively = true; // Flag to determine if content needs recursive parsing

        // --- 2. Determine outer match type and content ---
        if (match[2] && match[1] === '`') { // Code
            styleToApply = CODE_STYLE; contentToParse = match[2];
            processRecursively = false; // Don't parse inside code

        } else if (match[4]) { // Bold + Italic
            styleToApply = BOLD_ITALIC_STYLE; contentToParse = match[4];
        } else if (match[6]) { // Bold
            styleToApply = BOLD_STYLE; contentToParse = match[6];
        } else if (match[8]) { // Italic
            styleToApply = ITALIC_STYLE; contentToParse = match[8];
        } else if (match[10]) { // Strikethrough
            styleToApply = STRIKETHROUGH_STYLE; contentToParse = match[10];
        } else if (match[12]) { // <u> tag -> Apply Maroon
            styleToApply = MAROON_STYLE; contentToParse = match[12];
        } else if (match[15]) { // --- NEW: Square Brackets (Group 15 is content) ---
            styleToApply = DARK_BLUE_STYLE; contentToParse = match[15];
        } else if (match[17]) { // Plain Text Section (Group 17)
             finalSegments.push({ text: match[17], style: null });
             processRecursively = false; // No content to parse
        } else if (match[18]) { // Single Markdown/HTML/Bracket Character (Group 18)
             finalSegments.push({ text: match[18], style: null });
             processRecursively = false; // No content to parse
        } else { // Fallback
             finalSegments.push({ text: fullMatch, style: null });
             processRecursively = false;
        }

        // --- 3. Process content (recursively if needed) ---
        if (processRecursively && styleToApply) {
            processNestedContent(contentToParse, styleToApply, finalSegments);
        } else if (!processRecursively && styleToApply) {
            // Handle non-recursive cases like code blocks
             finalSegments.push({ text: contentToParse, style: styleToApply });
        }
        // If !processRecursively and !styleToApply, it was handled directly (plain text, single char)

        currentIndex = matchIndex + fullMatch.length; // Move index past the matched segment
    }

    // --- 4. Add any remaining plain text after the last match ---
    if (currentIndex < textLine.length) {
        finalSegments.push({ text: textLine.substring(currentIndex), style: null });
    }

    return finalSegments;
}

/**
 * Helper function to recursively parse content and apply an outer style.
 * (Corrected: Removed font size merging for inline styles)
 * @param {string} content The inner content string to parse.
 * @param {RichTextStyle} outerStyle The style to merge onto the parsed content's segments.
 * @param {Array} targetSegments The array to push the results into.
 */
function processNestedContent(content, outerStyle, targetSegments) {
    const innerSegments = parseInlineMarkdown(content); // Recursive call
    for (const segment of innerSegments) {
        const builder = SpreadsheetApp.newTextStyle();
        const existingStyle = segment.style;

        // Merge outerStyle properties onto existingStyle
        builder.setBold((outerStyle.isBold() ?? existingStyle?.isBold()) ?? false);
        builder.setItalic((outerStyle.isItalic() ?? existingStyle?.isItalic()) ?? false);
        builder.setStrikethrough((outerStyle.isStrikethrough() ?? existingStyle?.isStrikethrough()) ?? false);
        builder.setUnderline((outerStyle.isUnderline() ?? existingStyle?.isUnderline()) ?? false); // Keep just in case

        // Use the null-safe logic for properties that can be null
        builder.setForegroundColor(outerStyle.getForegroundColor() ?? existingStyle?.getForegroundColor());
        builder.setFontFamily(outerStyle.getFontFamily() ?? existingStyle?.getFontFamily());

        // --- REMOVED this line to prevent the error and let base style control font size ---
        // builder.setFontSize(outerStyle.getFontSize() ?? existingStyle?.getFontSize());

        const mergedStyle = builder.build();
        targetSegments.push({ text: segment.text, style: mergedStyle });
    }
}

/**
 * Shows an HTML Service dialog for pasting Markdown Tables.
 */
function showMarkdownTableDialog() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('MarkdownTableDialog')
      .setWidth(450)
      .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Paste Markdown Table');
}

/**
 * Processes Markdown Table text pasted from the HTML dialog.
 * Skips header and separator lines.
 * Splits data rows and cells, applies inline formatting, sets top border,
 * header background, text wrap, and vertical alignment.
 * @param {string} markdownTableText The raw Markdown table text.
 */
function processPastedMarkdownTable(markdownTableText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const activeCell = ss.getActiveRange()?.getCell(1, 1);

  if (!activeCell) { throw new Error("No active cell selected..."); }
  if (!markdownTableText || markdownTableText.trim() === "") { /* ... handle empty ... */ return; }

  const startRow = activeCell.getRow();
  const startCol = activeCell.getColumn();
  let maxCols = 0;
  const lines = markdownTableText.trim().split('\n');
  const rowsToPaste = [];
  let isFirstRealRow = true;

  // --- Loop through lines to process rows (same logic as before) ---
  for (const line of lines) {
    const trimmedLine = line.trim();
    if (trimmedLine === "" || trimmedLine.replace(/\|/g, "").match(/^[- :]+$/)) { continue; } // Skip separator/empty

    if (trimmedLine.includes('|')) {
      if (isFirstRealRow) { // Skip header
        isFirstRealRow = false;
        if (DEBUG_MODE) Logger.log(`Skipping header line: "${trimmedLine}"`);
        continue;
      }
      // Process DATA row
      const cellsRaw = trimmedLine.replace(/^\||\|$/g, '').split('|');
      const richTextCellsForRow = [];
      maxCols = Math.max(maxCols, cellsRaw.length);
      for (const cellRaw of cellsRaw) {
        const cellMarkdown = cellRaw.trim();
        const cellRichText = convertMarkdownCellToRichText(cellMarkdown);
        richTextCellsForRow.push(cellRichText);
      }
      rowsToPaste.push(richTextCellsForRow);
    }
  } // End line processing loop

  // --- Paste the processed DATA rows onto the sheet ---
  if (rowsToPaste.length > 0 && maxCols > 0) {
    try {
      const paddedRows = rowsToPaste.map(row => { /* ... same padding logic ... */
         while (row.length < maxCols) { row.push(SpreadsheetApp.newRichTextValue().setText("").build()); }
         return row;
      });

      const targetRange = sheet.getRange(startRow, startCol, paddedRows.length, maxCols);
      targetRange.clearContent().clearFormat();
      targetRange.setRichTextValues(paddedRows);
      if (DEBUG_MODE) Logger.log(`Pasted ${paddedRows.length} DATA rows and ${maxCols} columns.`);

      // --- Apply Top Border ---
      targetRange.setBorder( true, null, null, null, null, null, TABLE_TOP_BORDER_COLOR, TABLE_TOP_BORDER_STYLE );
      if (DEBUG_MODE) Logger.log(`Applied top border to range: ${targetRange.getA1Notation()}`);

      // --- Apply Background Color to FIRST data row ---
      if (paddedRows.length > 0) {
          const headerRange = sheet.getRange(startRow, startCol, 1, maxCols);
          headerRange.setBackground(TABLE_HEADER_BACKGROUND_COLOR);
          if (DEBUG_MODE) Logger.log(`Applied header background color to range: ${headerRange.getA1Notation()}`);
      }

      // --- NEW: Apply Text Wrap and Vertical Alignment to the entire data range ---
      targetRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      targetRange.setVerticalAlignment("middle"); // Options: "top", "middle", "bottom"
      if (DEBUG_MODE) Logger.log(`Applied text wrap and middle vertical alignment to range: ${targetRange.getA1Notation()}`);
      // --- End NEW ---

    } catch (error) {
      Logger.log(`Error pasting table data or setting formatting: ${error}\n${error.stack}`);
      throw new Error(`Failed to paste table or set formatting: ${error.message}`);
    }
  } else {
     if (DEBUG_MODE) Logger.log("No valid data rows found to paste.");
  }
}


/**
 * Converts Markdown within a SINGLE table cell to a RichTextValue object.
 * Reuses the parseInlineMarkdown logic.
 * @param {string} cellMarkdown The Markdown content of a single cell.
 * @return {RichTextValue} The formatted RichTextValue for the cell.
 */
function convertMarkdownCellToRichText(cellMarkdown) {
  const builder = SpreadsheetApp.newRichTextValue();
  if (!cellMarkdown) {
    return builder.setText("").build();
  }

  const segments = parseInlineMarkdown(cellMarkdown); // Use the existing inline parser
  let cellText = "";
  const styleRuns = []; // {start, end, style} relative to cellText

  for (const segment of segments) {
    const start = cellText.length;
    cellText += segment.text;
    const end = cellText.length;
    if (segment.style && start < end) {
       // Use the globally defined styles (includes Italic color etc)
      styleRuns.push({ start: start, end: end, style: segment.style });
    }
  }

  builder.setText(cellText);
  // Apply DEFAULT_STYLE first to the whole cell? Optional, depends on desired base format.
  // builder.setTextStyle(0, cellText.length, DEFAULT_STYLE);

  // Apply specific inline styles collected
  for (const run of styleRuns) {
     try {
         builder.setTextStyle(run.start, run.end, run.style);
     } catch (e) {
         Logger.log(`Error applying style to cell run [${run.start}, ${run.end}): ${e}`);
     }
  }

  return builder.build();
}

// ==========================================================
// --- NEW: Find, Replace & Format Tool Functions ---
// ==========================================================

/**
 * Shows a sidebar in the spreadsheet for the Find & Format tool.
 * This is called from the new menu item.
 */
function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar')
      .evaluate()
      .setTitle('Find, Replace & Format')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * The main function that finds text and applies formatting and optional replacement.
 *
 * This version FIXES the bug where subsequent runs would wipe out previous formatting.
 * It separates logic for "Format Only" (preserves existing formats) and
 * "Replace & Format" (rebuilds the cell, which may clear other formats in that cell).
 *
 * @param {object} options The user-defined options from the sidebar.
 * @returns {string} A success message.
 */
function processFindAndFormat(options) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    
    // Determine the range to search
    let range;
    if (options.scope === 'selection') {
      range = SpreadsheetApp.getActiveRange();
      if (!range) {
        throw new Error("Please select a range of cells first.");
      }
    } else {
      range = sheet.getDataRange();
    }

    // Get the search terms
    let searchTerms = [];
    if (options.searchType === 'list') {
      searchTerms = options.searchTerms.split(',').map(term => term.trim()).filter(Boolean);
    } else {
      searchTerms = [options.searchTerms];
    }
    
    if (searchTerms.length === 0) {
      throw new Error("Please provide one or more search terms.");
    }
    
    // Build the text style for the new format
    const styleBuilder = SpreadsheetApp.newTextStyle();
    if (options.isBold) styleBuilder.setBold(true);
    if (options.isItalic) styleBuilder.setItalic(true);
    if (options.fontColor) styleBuilder.setForegroundColor(options.fontColor);
    if (options.fontSize) styleBuilder.setFontSize(parseInt(options.fontSize, 10));
    const newStyle = styleBuilder.build();

    const richTextValues = range.getRichTextValues();

    // Loop through each cell
    for (let i = 0; i < richTextValues.length; i++) {
      for (let j = 0; j < richTextValues[i].length; j++) {
        const cellValue = richTextValues[i][j];
        const cellText = cellValue.getText();
        
        if (!cellText) {
          continue; // Skip empty cells
        }

        // --- LOGIC SPLIT: Based on whether user is replacing text or just formatting ---

        if (options.replaceText && options.replaceText.length > 0) {
          // --- CASE 1: REPLACE & FORMAT ---
          // This operation is "destructive" to other formats in the cell because indices change.
          // We rebuild the cell's rich text from scratch.
          let newText = cellText;
          for (const term of searchTerms) {
            const escapedTerm = escapeRegExp(term);
            const regex = options.exactMatch 
              ? new RegExp(`\\b${escapedTerm}\\b`, 'g') 
              : new RegExp(escapedTerm, 'g');
            newText = newText.replace(regex, options.replaceText);
          }
          
          // Create a new builder with the fully replaced text
          const replacedBuilder = SpreadsheetApp.newRichTextValue().setText(newText);
          
          // Find all instances of the replacement text and apply the new style
          const replacedTermRegex = new RegExp(escapeRegExp(options.replaceText), 'g');
          let replacedMatch;
          while ((replacedMatch = replacedTermRegex.exec(newText)) !== null) {
              const start = replacedMatch.index;
              const end = start + replacedMatch[0].length;
              replacedBuilder.setTextStyle(start, end, newStyle);
          }
          richTextValues[i][j] = replacedBuilder.build(); // Update the value in our array

        } else {
          // --- CASE 2: FORMAT ONLY ---
          // This operation is "non-destructive". It preserves existing formatting.
          
          // THE KEY FIX: Create a builder by copying the existing rich text value.
          const richTextBuilder = cellValue.copy();
          
          for (const term of searchTerms) {
            const escapedTerm = escapeRegExp(term);
            const regex = options.exactMatch 
              ? new RegExp(`\\b${escapedTerm}\\b`, 'g') 
              : new RegExp(escapedTerm, 'g');
            
            let match;
            while ((match = regex.exec(cellText)) !== null) {
              const startIndex = match.index;
              const endIndex = startIndex + match[0].length;
              // Apply the new style on top of any existing styles
              richTextBuilder.setTextStyle(startIndex, endIndex, newStyle);
            }
          }
          richTextValues[i][j] = richTextBuilder.build(); // Update the value in our array
        }
      }
    }
    
    // Set the new rich text values back to the range in one single operation
    range.setRichTextValues(richTextValues);
    
    return 'Processing complete!';

  } catch (e) {
    Logger.log(e);
    throw new Error('Error: ' + e.message);
  }
}

/**
 * Helper function to escape special characters for use in a regular expression.
 */
function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}