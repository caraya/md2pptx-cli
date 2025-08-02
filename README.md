# md2pptx-cli

`md2pptx-cli` is a command-line tool that converts Markdown files into PowerPoint (`.pptx`) presentations. It leverages the power of `marked` for Markdown parsing and `PptxGenJS` for presentation generation, allowing you to create professional-looking slides directly from simple text files.

## Features

- **Slide Creation**: Each top-level heading (`# Heading`) creates a new slide.
- **Text Formatting**: Supports **bold**, *italic*, `` `code` ``, and [hyperlinks](https://www.example.com).
- **Lists**: Handles bulleted lists and task lists (`- [x]`).
- **Tables**: Converts standard Markdown tables into PowerPoint tables.
- **Images**: Embed images from local paths or URLs.
- **Two-Column Layout**: Use a horizontal rule (`---`) to split a slide's content area into two columns.
- **Speaker Notes**: Add speaker notes that are visible in presenter view but not on the slide itself.
- **Shapes**: Embed basic shapes like rectangles and ovals with custom styling.
- **Auto-Fitting Text**: Text automatically resizes to fit its container, preventing overflow.

## Installation

1. **Clone the repository:**

    ```bash
    git clone <repository-url>
    cd md2pptx-cli
    ```

2. **Install dependencies:**

    ```bash
    npm install
    ```

3. **Build the tool:**

    ```bash
    npm run build
    ```

4. **(Optional) Link for global use:**
    To make the `md2pptx` command available anywhere on your system, run:

    ```bash
    npm link
    ```

## Usage

Run the tool from your terminal, providing an input Markdown file. The output file will be automatically named (e.g., `input.md` -> `input.pptx`) unless you specify a different name.

### Basic Usage

```bash
md2pptx <input.md> [output.pptx]
```

## Examples

Convert `myslides.md` to `myslides.pptx`:

```bash
md2pptx myslides.md
```

Convert `report.md` to `presentation.pptx`:

```bash
md2pptx report.md presentation.pptx
```

## Markdown Syntax Guide

Here is a guide to the supported Markdown syntax for creating your presentation.

### Slides and Titles

Each Level 1 heading creates a new slide with that heading as its title.

```markdown
# This is the Title of Slide 1

Content for the first slide.

# This is the Title of Slide 2

Content for the second slide.
```

### Two-Column Layout

To treat the entire content area of a slide as a two-column layout, use a single horizontal rule (`---`). All content before the `---` will appear in the left column, and all content after it will appear in the right column.

```markdown
# Two-Column Slide

This is the left column.
- A bullet point
- *More text*

---

This is the right column.

!shape[rect]({"w":4, "h":2, "fill":{"color":"FF0000"}})
```

### Speaker Notes

Use blockquote syntax with the prefix Note: to add speaker notes.

```bash
> Note: This is a speaker note.
```

### Inline Formatting

Bold: `**bold text**`

Italic: `*italic text*`

Code: `\`inline code\``

Hyperlink: `[link text](https://www.example.com)`

### Images

Use the standard Markdown image syntax.

```markdown
![Alt text for the image](https://placehold.co/400x200/cccccc/ffffff?text=Image)
```

### Tables

Create tables using standard Markdown table syntax.

```markdown
| Header 1 | Header 2 |
|----------|----------|
| Cell 1-1 | Cell 1-2 |
| Cell 2-1 | Cell 2-2 |
```

### Shapes

Embed shapes using a special !shape syntax. The shape type is specified in brackets, and the options (following the PptxGenJS API) are provided as a JSON object in parentheses. Do not include x or y coordinates; they are calculated automatically.

Syntax: `!shape[<shapeType>]({<options>})`

Available Shape Types: rect, oval, ellipse, line, triangle

Example:

```markdown
# Slide with Shapes

!shape[rect]({"w":4, "h":2, "fill":{"color":"FF0000"}})

!shape[oval]({"w":3, "h":1.5, "fill":{"color":"0000FF"}})
```

## Limitations

**Layout**: To ensure stability and prevent file corruption, a slide can be either single-column or two-column. Mixing full-width content with two-column content on the same slide is not supported. If you need full-width text before a two-column section, please place it on a separate slide
