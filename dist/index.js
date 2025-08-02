#!/usr/bin/env node
import fs from 'fs/promises';
import { marked } from 'marked';
import PptxGenJS from 'pptxgenjs';
/** Print usage and exit. */
function usage() {
    console.error('Usage: md2pptx <input.md> [output.pptx]');
    process.exit(1);
}
/**
 * Turn raw markdown into an array of SlideData.
 */
function parseSlides(md) {
    const tokens = marked.lexer(md, { gfm: true });
    const slides = [];
    let current = null;
    for (const token of tokens) {
        switch (token.type) {
            case 'heading':
                if (token.depth === 1) {
                    if (current)
                        slides.push(current);
                    current = { title: token.text, elements: [] };
                }
                break;
            case 'paragraph':
                if (!current)
                    break;
                const noteMatch = token.text.match(/^>\s*Note:\s*(.*)/);
                if (noteMatch) {
                    current.notes = noteMatch[1].trim();
                    break;
                }
                const txt = token.text.trim();
                const img = txt.match(/^!\[([^\]]*)\]\(([^)]+)\)$/);
                const shapeMatch = txt.match(/^!shape\[(.*)\]\((.*)\)$/);
                if (img) {
                    current.elements.push({ type: 'image', alt: img[1], href: img[2] });
                }
                else if (shapeMatch) {
                    try {
                        const shapeType = shapeMatch[1];
                        const options = JSON.parse(shapeMatch[2]);
                        current.elements.push({ type: 'shape', shape: shapeType, options });
                    }
                    catch (e) {
                        console.error("Invalid shape syntax:", txt);
                    }
                }
                else {
                    const elements = [];
                    const linkRegex = /\[([^\]]+)\]\(([^)]+)\)/g;
                    let lastIndex = 0;
                    let match;
                    while ((match = linkRegex.exec(txt)) !== null) {
                        if (match.index > lastIndex) {
                            elements.push({ type: 'text', text: txt.substring(lastIndex, match.index) });
                        }
                        elements.push({ type: 'text', text: match[1], options: { hyperlink: { url: match[2] } } });
                        lastIndex = match.index + match[0].length;
                    }
                    if (lastIndex < txt.length) {
                        elements.push({ type: 'text', text: txt.substring(lastIndex) });
                    }
                    current.elements.push(...elements);
                }
                break;
            case 'list':
                if (!current)
                    break;
                for (const item of token.items) {
                    if (item.task) {
                        current.elements.push({ type: 'task', text: item.text, checked: item.checked });
                    }
                    else {
                        current.elements.push({ type: 'bullet', text: item.text });
                    }
                }
                break;
            case 'table':
                if (!current)
                    break;
                const headers = token.header.map(cell => cell.text);
                const rows = token.rows.map(row => row.map(cell => cell.text));
                current.elements.push({ type: 'table', headers, rows });
                break;
            case 'hr':
                if (current)
                    current.elements.push({ type: 'columnBreak' });
                break;
            default:
                break;
        }
    }
    if (current)
        slides.push(current);
    return slides;
}
/**
 * Render a list of SlideElements into one text box or sequence of images.
 */
function renderElements(slide, elems, x, y, w) {
    let yOffset = y;
    const textRuns = [];
    const flushText = () => {
        if (!textRuns.length)
            return;
        slide.addText(textRuns, { x, y: yOffset, w, fontSize: 18, lineSpacing: 20 });
        yOffset += textRuns.length * 0.3;
        textRuns.length = 0;
    };
    for (const el of elems) {
        if (el.type === 'text') {
            textRuns.push({ text: el.text, options: el.options });
        }
        else if (el.type === 'bullet') {
            textRuns.push({ text: el.text, options: { bullet: true } });
        }
        else if (el.type === 'task') {
            const box = el.checked ? '☑️' : '☐';
            textRuns.push({ text: `${box} ${el.text}` });
        }
        else if (el.type === 'image') {
            flushText();
            slide.addImage({ path: el.href, x, y: yOffset, w, h: 3 });
            yOffset += 3.2;
        }
        else if (el.type === 'table') {
            flushText();
            // FIX: `addTable` expects an array of TableRow objects, where each row
            // is an array of TableCell objects, not just strings.
            const tableData = [el.headers, ...el.rows].map(row => row.map(text => ({ text })));
            slide.addTable(tableData, { x, y: yOffset, w });
            yOffset += el.rows.length * 0.5 + 0.5;
        }
        else if (el.type === 'shape') {
            flushText();
            slide.addShape(el.shape, { ...el.options, x, y: yOffset, w });
            yOffset += (el.options.h || 1) + 0.2;
        }
    }
    flushText();
}
/**
 * Build the PPTX, doing single or two-column layout per slide.
 */
async function buildPptx(slides, outPath) {
    const pptx = new PptxGenJS();
    for (const { title, elements, notes } of slides) {
        const slide = pptx.addSlide();
        slide.addText(title, { x: 0.5, y: 0.3, fontSize: 32, bold: true });
        if (notes) {
            slide.addNotes(notes);
        }
        const idx = elements.findIndex((e) => e.type === 'columnBreak');
        if (idx >= 0) {
            const col1 = elements.slice(0, idx);
            const col2 = elements.slice(idx + 1);
            renderElements(slide, col1, 0.5, 1.2, 4.5);
            renderElements(slide, col2, 5.0, 1.2, 4.5);
        }
        else {
            renderElements(slide, elements, 0.5, 1.2, 9.0);
        }
    }
    await pptx.writeFile({ fileName: outPath });
}
/** CLI entrypoint */
async function main() {
    const [, , inFile, maybeOut] = process.argv;
    if (!inFile) {
        usage();
    }
    const outFile = maybeOut || inFile.replace(/\.md$/i, '.pptx');
    try {
        const md = await fs.readFile(inFile, 'utf-8');
        const slides = parseSlides(md);
        await buildPptx(slides, outFile);
        console.log(`✅ Written ${outFile}`);
    }
    catch (err) {
        console.error('Error:', err.message || err);
        process.exit(1);
    }
}
main();
//# sourceMappingURL=index.js.map