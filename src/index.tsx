#!/usr/bin/env tsx
import fs from 'fs/promises'
import { marked } from 'marked'
import PptxGenJS from 'pptxgenjs'

// This correctly infers the type of a slide object from the default export.
type Slide = ReturnType<InstanceType<typeof PptxGenJS>['addSlide']>;

// Define the ShapeType enum locally to avoid module resolution issues at runtime.
enum LocalShapeType {
    RECTANGLE = 'rect',
    // The PptxGenJS library expects the string 'ellipse' for ovals.
    OVAL = 'ellipse',
    ELLIPSE = 'ellipse',
    LINE = 'line',
    TRIANGLE = 'triangle',
}

/**
 * A map to convert simple names from Markdown to the official PptxGenJS shape types.
 */
const shapeNameMap: { [key: string]: LocalShapeType } = {
    rect: LocalShapeType.RECTANGLE,
    oval: LocalShapeType.OVAL,
    ellipse: LocalShapeType.ELLIPSE,
    line: LocalShapeType.LINE,
    triangle: LocalShapeType.TRIANGLE,
};

// Represents a block of formatted text.
type TextBlock = {
    type: 'textBlock';
    runs: PptxGenJS.TextProps[];
};

/**
 * One piece of slide content.
 */
type SlideElement =
  | TextBlock
  | { type: 'image'; href: string; alt: string }
  | { type: 'table'; headers: string[]; rows: string[][] }
  | { type: 'shape'; shape: LocalShapeType; options: PptxGenJS.ShapeProps }
  | { type: 'columnBreak' }

/** A slide’s title and its elements in reading order. */
interface SlideData {
  title: string
  elements: SlideElement[]
  notes?: string
}

/** Print usage and exit. */
function usage(): never {
  console.error('Usage: md2pptx <input.md> [output.pptx]')
  process.exit(1)
}

/**
 * Recursively parses tokens from `marked` to handle inline formatting like bold, italics, and links.
 * This function has been rewritten to correctly handle nested formatting.
 */
function parseTextRuns(tokens: marked.Token[], parentOptions: PptxGenJS.TextPropsOptions = {}): PptxGenJS.TextProps[] {
    const runs: PptxGenJS.TextProps[] = [];

    for (const token of tokens) {
        const currentOptions = { ...parentOptions };

        switch (token.type) {
            case 'strong':
                runs.push(...parseTextRuns(token.tokens, { ...currentOptions, bold: true }));
                break;
            case 'em':
                runs.push(...parseTextRuns(token.tokens, { ...currentOptions, italic: true }));
                break;
            case 'codespan':
                runs.push({ text: ` ${token.text} `, options: { ...currentOptions, fontFace: 'Courier New' } });
                break;
            case 'link':
                const linkOptions = { ...currentOptions, hyperlink: { url: token.href, tooltip: token.title || token.href } };
                runs.push(...parseTextRuns(token.tokens, linkOptions));
                break;
            case 'text':
                // FIX: Add a type guard. The 'text' type can be ambiguous in `marked`.
                // We must check for the existence of the `tokens` property before accessing it.
                if ('tokens' in token && token.tokens) {
                     runs.push(...parseTextRuns(token.tokens, currentOptions));
                } else {
                    runs.push({ text: token.text, options: currentOptions });
                }
                break;
            default:
                if (token.raw) {
                    runs.push({ text: token.raw, options: currentOptions });
                }
                break;
        }
    }
    return runs;
}


/**
 * Turn raw markdown into an array of SlideData.
 */
function parseSlides(md: string): SlideData[] {
  const tokens = marked.lexer(md, { gfm: true })
  const slides: SlideData[] = []
  let current: SlideData | null = null

  for (const token of tokens) {
    switch (token.type) {
      case 'heading':
        if (token.depth === 1) {
          if (current) slides.push(current)
          current = { title: token.text, elements: [] }
        }
        break

      case 'paragraph':
        if (!current) break;
        const noteMatch = token.text.match(/^>\s*Note:\s*(.*)/);
        if (noteMatch) {
            current.notes = noteMatch[1].trim();
            break;
        }

        const txt = token.text.trim()
        const img = txt.match(/^!\[([^\]]*)\]\(([^)]+)\)$/)
        const shapeMatch = txt.match(/^!shape\[(.*)\]\((.*)\)$/);

        if (img) {
          current.elements.push({ type: 'image', alt: img[1], href: img[2] })
        } else if (shapeMatch) {
            try {
                const shapeName = shapeMatch[1];
                const shapeType = shapeNameMap[shapeName];
                if (!shapeType) {
                    console.error("[ERROR] Invalid shape name:", shapeName);
                    break;
                }
                const options = JSON.parse(shapeMatch[2]);
                current.elements.push({ type: 'shape', shape: shapeType, options });
            } catch(e) {
                console.error("[ERROR] Invalid shape syntax:", txt);
            }
        } else {
          const runs = parseTextRuns(token.tokens || []);
          if (runs.length > 0) {
              current.elements.push({ type: 'textBlock', runs });
          }
        }
        break

      case 'list':
        if (!current) break;
        // FIX: Cast the token to the correct type to avoid using `any`.
        const listToken = token as marked.Tokens.List;
        for (const item of listToken.items) {
            const runs = parseTextRuns(item.tokens);
            if (runs.length > 0) {
                const textBlock: TextBlock = { type: 'textBlock', runs: [] };
                if (item.task) {
                    const box = item.checked ? '☑️ ' : '☐ ';
                    if (runs[0].text) {
                        runs[0].text = box + runs[0].text;
                    } else {
                        runs.unshift({ text: box });
                    }
                } else {
                    if (!runs[0].options) runs[0].options = {};
                    runs[0].options.bullet = true;
                }
                textBlock.runs.push(...runs);
                current.elements.push(textBlock);
            }
        }
        break

      case 'table':
        if (!current) break;
        const headers = token.header.map(cell => cell.text);
        const rows = token.rows.map(row => row.map(cell => cell.text));
        current.elements.push({ type: 'table', headers, rows });
        break;

      case 'hr':
        if (current) {
            current.elements.push({ type: 'columnBreak' });
        }
        break

      default:
        break
    }
  }

  if (current) slides.push(current)
  return slides
}

/**
 * Renders a list of elements into a slide column and returns the final Y-offset.
 */
function renderElements(
  slide: Slide,
  elems: SlideElement[],
  x: number,
  y: number,
  w: number
): number {
  let yOffset = y;
  const slideHeight = 5.625; // Standard 16:9 slide height
  const bottomMargin = 0.25;

  for (const el of elems) {
    const remainingHeight = slideHeight - yOffset - bottomMargin;
    if (remainingHeight <= 0.1) {
        break;
    }

    if (el.type === 'textBlock') {
        const estimatedLines = el.runs.reduce((acc, run) => acc + (run.options?.breakLine ? 1 : 0), 1);
        const estimatedHeight = Math.min(remainingHeight, estimatedLines * 0.5);
        slide.addText(el.runs, { x, y: yOffset, w, h: estimatedHeight, autoFit: true, fontSize: 18, lineSpacing: 24 });
        yOffset += estimatedHeight + 0.1;
    } else if (el.type === 'image') {
      slide.addImage({ path: el.href, x, y: yOffset, w, h: 3 });
      yOffset += 3.2;
    } else if (el.type === 'table') {
      const tableData = [el.headers, ...el.rows].map(row => row.map(text => ({ text })));
      slide.addTable(tableData, { x, y: yOffset, w, autoPage: true });
      yOffset += el.rows.length * 0.5 + 0.5;
    } else if (el.type === 'shape') {
      const finalOptions = { x, w, ...el.options, y: yOffset };
      slide.addShape(el.shape as unknown as PptxGenJS.ShapeType, finalOptions);
      yOffset += (finalOptions.h as number || 1) + 0.2;
    }
  }
  return yOffset;
}

/**
 * Build the PPTX, doing single or two-column layout per slide.
 */
async function buildPptx(slides: SlideData[], outPath: string) {
  const pptx = new PptxGenJS()

  for (const { title, elements, notes } of slides) {
    const slide = pptx.addSlide()
    slide.addText(title, { x: 0.5, y: 0.3, w: 9.0, h: 0.75, fontSize: 32, bold: true, autoFit: true })

    if (notes) {
        slide.addNotes(notes);
    }

    const breakIdx = elements.findIndex((e) => e.type === 'columnBreak');
    let yPos = 1.2;

    if (breakIdx === -1) {
        // No column break, render all elements as a single, full-width column.
        renderElements(slide, elements, 0.5, yPos, 9.0);
    } else {
        // A column break exists. The entire slide is treated as a two-column layout.
        const col1Elements = elements.slice(0, breakIdx);
        const col2Elements = elements.slice(breakIdx + 1);

        renderElements(slide, col1Elements, 0.5, yPos, 4.5);
        renderElements(slide, col2Elements, 5.2, yPos, 4.5);
    }
  }

  await pptx.writeFile({ fileName: outPath })
}

/** CLI entrypoint */
async function main() {
  const [,, inFile, maybeOut] = process.argv
  if (!inFile) {
    usage()
  }
  const outFile = maybeOut || inFile.replace(/\.md$/i, '.pptx')

  try {
    const md = await fs.readFile(inFile, 'utf-8')
    const slides = parseSlides(md)
    await buildPptx(slides, outFile)
    console.log(`\n✅ Written ${outFile}`)
  } catch (err: any) {
    console.error('[FATAL ERROR]', err.message || err)
    process.exit(1)
  }
}

main()
