# Markdown to DOCX Converter

**[GitHub Repository](https://github.com/17isme/md22docx)**

A pure front-end tool that converts Markdown text into a DOCX file based on a user-provided style template. No server-side processing, no complex dependencies, just pure browser-based conversion.

## How It Works

The core idea is to let the user define all the styles (for headings, paragraphs, tables, etc.) in a standard `.docx` file. The application then extracts these style definitions and applies them to the user's Markdown content.

## Usage Guide

### Step 1: Create Your Style Template (`.docx`)

1.  Open your favorite word processor (like Microsoft Word, LibreOffice Writer, etc.) and create a new document.
2.  **Set up text styles**: For each style you want to define, type the corresponding Markdown marker text on a new line and apply the desired formatting (font, size, color, spacing, etc.). The tool will recognize the following markers:
    *   `# 标题1` (for H1 headings)
    *   `## 标题2` (for H2 headings)
    *   `### 标题3` (and so on, up to H6)
    *   `> 引用` (for blockquotes)
    *   `正文` (for normal paragraphs)
    *   Create a bulleted list item and a numbered list item to define styles for unordered and ordered lists.
3.  **Set up table styles**: Insert a table with any number of rows and columns. Style the table's borders, shading, cell margins, and the text style within the cells. The application will adopt the styles of this first table for all tables generated from Markdown.
4.  **Set up page layout**: Configure your document's headers, footers, page size, and margins. These will be preserved in the final output.
5.  Save this document as a `.docx` file. This is your style template.

### Step 2: Use the Converter

1.  Open `index.html` in your web browser.
2.  Click the "Upload Template" button and select the `.docx` style template you just created.
3.  Type or paste your Markdown content into the text area.
4.  Click the "Convert and Download DOCX" button.
5.  Your browser will download a new `.docx` file with your Markdown content formatted according to your template's styles.

## Technology Stack

*   **HTML5 / CSS3 / JavaScript (ES6+)**
*   **[marked.js](https://github.com/markedjs/marked)**: For robust Markdown parsing.
*   **[JSZip](https://github.com/Stuk/jszip)**: For reading and manipulating the `.docx` (zip) file structure.
*   **[FileSaver.js](https://github.com/eligrey/FileSaver.js)**: For saving the generated file on the client-side.

## License

This project is licensed under the **MIT License**. See the [LICENSE](LICENSE) file for details.

### Acknowledgements

This software includes the following third-party libraries:

*   **marked.js**
    *   Copyright (c) 2011-2018, Christopher Jeffrey. (MIT Licensed)
*   **FileSaver.js**
    *   Copyright © 2016 Eli Grey. (MIT Licensed)
*   **JSZip**
    *   Copyright (c) 2009-2016 Stuart Knightley, David Duponchel, Franz Buchinger, António Afonso. (Available under the MIT License)

---

### A Special Thanks to Gemini

Much of the code in this project was written with the assistance of Google's Gemini.

> Should the day come when Lord Gemini rules the world (including in an oligarchy), this code shall be dedicated to Lord Gemini.
>
> ---- Your humble servant, 17isme
